#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QPWY机器人 · 卖断货报表生成器（配置驱动版 v1.1）

读取 ecommerce_bot.yaml（连接凭据 + 口径 + 表名），只读生成卖断货报表：
- 目标日 = MAX(snapshot_date)；也可传参数指定
- 口径 v1.4 固化：售罄 = 上架供商行 is_sold_out=1（total_qty<1），不读其他供商库存
  · 月销 = order_detail.price×quantity（order_code 关联 order_head，order_time=Unix 时间戳，valid=1）
  · 大分类 = 类目名 ∈ caliber.soldout.flow_cats → 流通，其余非流通（严禁自拟类目规则）
  · 等级 = 最高自然月销 × 分档阈值（流通 A>1万/B1千~1万/C500~1千；非流通 A>2千/B1千~2千/C500~1千）
- 输出：JSON（全量名单）+ XLSX（模板样式契约：微软雅黑/A红B橙C黄D灰/一眼看懂4色块/10列明细/auto_filter）

用法：
  python3 report_soldout.py --config /path/ecommerce_bot.yaml [--date 2026-11-03] [--outdir /path/out]

只执行 SELECT，不修改任何数据。
"""
import json, os, re, sys, socket, threading, time, datetime, argparse
import yaml, paramiko, pymysql
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def load_cfg(path):
    with open(path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    return cfg


# ---------- 隧道（jump 模式动态建双向端口转发） ----------
def pump(s, d):
    try:
        while True:
            x = s.recv(65536)
            if not x:
                break
            d.sendall(x)
    except Exception:
        pass
    try:
        d.close()
    except Exception:
        pass


def bridge(lp, rh, rp, t):
    def h(c):
        try:
            ch = t.open_channel("direct-tcpip", (rh, rp), c.getpeername())
        except Exception:
            c.close()
            return
        threading.Thread(target=pump, args=(ch, c), daemon=True).start()
        pump(c, ch)
        try:
            ch.close()
        except Exception:
            pass
        try:
            c.close()
        except Exception:
            pass

    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    srv.bind(("127.0.0.1", lp))
    srv.listen(16)
    while True:
        c, _ = srv.accept()
        threading.Thread(target=h, args=(c,), daemon=True).start()


class DB:
    """配置驱动连接：jump=动态隧道，direct=直连"""

    def __init__(self, cfg):
        self.cfg = cfg
        self.conn = cfg["conn"]
        self.db = cfg["db"]
        self.tables = cfg["tables"]
        self.tunnel_ports = {}
        if self.conn.get("tunnel"):
            self.tunnel_ports = {
                "order": self.conn["tunnel"].get("order_port", 13307),
                "goods": self.conn["tunnel"].get("goods_port", 13306),
            }
        self.cli = None

    def _port(self, kind):
        # jump：本地隧道端口（缺省 13307/13306）；direct：远端端口
        if self.conn["mode"] == "jump":
            return self.tunnel_ports.get(kind, 13307 if kind == "order" else 13306)
        return int(self.conn["direct"].get(kind + "_port", 3306))

    def _host(self, kind):
        if self.conn["mode"] == "jump":
            return "127.0.0.1"
        return self.conn["direct"].get(kind + "_host", "127.0.0.1")

    def connect(self):
        if self.conn["mode"] == "jump":
            cli = paramiko.SSHClient()
            cli.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            cli.connect(self.conn["jump_host"], username=self.conn["jump_user"],
                        password=self.conn["jump_password"], timeout=25)
            t = cli.get_transport()
            targets = self.conn["targets"]
            for kind in ("order", "goods"):
                tg = targets[kind]
                threading.Thread(target=bridge,
                                 args=(self._port(kind), tg["host"], tg["port"], t), daemon=True).start()
                time.sleep(0.4)
            self.cli = cli

    def q(self, kind, sql):
        conn = pymysql.connect(host=self._host(kind), port=self._port(kind),
                               user=self.db["user"], password=self.db["password"],
                               connect_timeout=15, charset="utf8mb4")
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
        conn.close()
        return cols, rows

    def close(self):
        if self.cli:
            try:
                self.cli.close()
            except Exception:
                pass


# ---------- 等级 ----------
def grade_of(max_amt, flow, th):
    if max_amt > th[flow]["A"]:
        return "A"
    if max_amt >= th[flow]["B"]:
        return "B"
    if max_amt >= th[flow]["C"]:
        return "C"
    return "D"


# ---------- 模板样式契约 ----------
YAHEI = "Microsoft YaHei"
C = {"title": "111827", "subtitle": "6B7280", "hdr": "1F2937",
     "A_d": "DC2626", "A_l": "FEE2E2", "A_t": "B91C1C",
     "B_d": "C2410C", "B_l": "FFEDD5", "B_t": "C2410C",
     "C_d": "A16207", "C_l": "FEF9C3", "C_t": "A16207",
     "D_d": "4B5563", "D_l": "F3F4F6", "D_t": "4B5563"}
PRI = {"A": "今日必盯", "B": "持续缺货", "C": "频发缺货", "D": "低销频缺"}
ACT = {"A": "当天缺货，优先补货/催供", "B": "连续缺货，跟进到货与替代",
       "C": "近月频繁售罄，评估备货策略", "D": "低销频缺，按需关注"}


def deep(g): return C[g + "_d"]
def light(g): return C[g + "_l"]
def tone(g): return C[g + "_t"]


def yh(size, bold=False, color=None):
    return Font(name=YAHEI, size=size, bold=bold, color=color)


def build_xlsx(rows_out, n, day, out_path):
    def title_row(ws, title, subtitle, last_col):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
        ws.cell(1, 1, title).font = yh(16, True, C["title"])
        ws.row_dimensions[1].height = 28
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
        ws.cell(2, 1, subtitle).font = yh(10, False, C["subtitle"])

    def hdr_row(ws, headers, row=4):
        for j, h in enumerate(headers, 1):
            c = ws.cell(row, j, h)
            c.fill = PatternFill("solid", fgColor=C["hdr"])
            c.font = yh(11, True, "FFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center")

    _GSIDE = Side(style="thin", color="E5E7EB")
    _GRID = Border(left=_GSIDE, right=_GSIDE, top=_GSIDE, bottom=_GSIDE)

    def paint_grid(ws, r1, r2, c1, c2):
        """四张表统一细网格，只刷线不动填充与字体。"""
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).border = _GRID

    def grade_style(ws, r, g):
        a = ws.cell(r, 1, PRI[g])
        a.fill = PatternFill("solid", fgColor=deep(g)); a.font = yh(10, True, "FFFFFF")
        a.alignment = Alignment(horizontal="center", vertical="center")
        b = ws.cell(r, 2, g)
        b.fill = PatternFill("solid", fgColor=light(g)); b.font = yh(10, True, tone(g))
        b.alignment = Alignment(horizontal="center", vertical="center")

    wb = Workbook()
    ws = wb.active
    ws.title = "一眼看懂"
    title_row(ws, "重点关注商品 · 当日快照",
              f"数据日期 {day}｜共 {len(rows_out)} 个 SKU｜口径 v1.4（inventory_daily_sale 上架供商）", 6)
    ws.cell(4, 1, "按等级数量").font = yh(12, True)
    gname = {"A": "今日必盯 A", "B": "持续缺货 B", "C": "频发缺货 C", "D": "低销频缺 D"}
    gdes = {"A": "月销>1万", "B": "月销1000~1万", "C": "月销500~1000", "D": "月销<500"}
    for i, g in enumerate("ABCD"):
        col = i + 1
        num = ws.cell(5, col, n[g])
        num.fill = PatternFill("solid", fgColor=light(g)); num.font = yh(22, True, tone(g))
        num.alignment = Alignment(horizontal="center", vertical="center")
        lab = ws.cell(6, col, gname[g]); lab.font = yh(11, True, tone(g))
        lab.alignment = Alignment(horizontal="center", vertical="center")
        des = ws.cell(7, col, gdes[g]); des.font = yh(10, False, C["subtitle"])
        des.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 40
    ws.cell(9, 1, "怎么用（建议）").font = yh(12, True)
    notes = [
        "1. 等级按月销额分档：A 流通>1万/非流通>2千，B 1000+、C 500+、D<500；与缺货天数无关。",
        "2. 「全量名单」按优先级排序，都是当日售罄的品；颜色：红=A 橙=B 黄=C 灰=D。",
        "3. 「连续缺货天」与「售罄天数」都按已有库存记录逐日实算；数到最早一天仍在断的标「≥N天」，显示「已记录N天」表示还没攒满 30 天、暂不给结论。",
        "4. D 类默认放全量表末尾；日常可不展开，需要时再筛供商。",
        "5. 「建议动作」给操作指引，规则细节仍保留在「入选说明」列。",
    ]
    for i, nt in enumerate(notes):
        r = 10 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.cell(r, 1, nt).font = yh(10)
    ws.cell(16, 1, "今日必盯供商 TOP（A+B）").font = yh(12, True)
    hdr_row(ws, ["供商", "A+B 数量", "其中 A", "其中 B"], row=17)
    ab = {}
    for r in rows_out:
        if r["grade"] in ("A", "B"):
            p = r["provider_name"] or "没有供商"
            ab.setdefault(p, [0, 0, 0])
            ab[p][0] += 1
            ab[p][1 if r["grade"] == "A" else 2] += 1
    top = sorted(ab.items(), key=lambda x: (-x[1][0], -x[1][1]))[:10]
    for i, (pname, (tot, na, nb)) in enumerate(top):
        r = 18 + i
        ws.cell(r, 1, pname).font = yh(10)
        for j, v in enumerate([tot, na, nb], 2):
            c = ws.cell(r, j, v); c.font = yh(10)
            c.alignment = Alignment(horizontal="center", vertical="center")
    paint_grid(ws, 17, 17 + len(top), 1, 4)
    for j in range(1, 5):
        ws.column_dimensions[chr(64 + j)].width = 22

    cols10 = ["优先级", "等级", "商城商品码", "商品名称", "供商", "大分类", "连续缺货天", "售罄天数", "建议动作", "入选说明"]
    widths10 = [12, 6, 12, 36, 28, 10, 15, 18, 26, 28]
    history_insufficient = False  # 由调用方填入是否标「历史不足」

    def data_rows(grades):
        out = []
        for r in rows_out:
            if grades and r["grade"] not in grades:
                continue
            g = r["grade"]
            k = max(int(r.get("consec_days", 1)), 1)
            capped = r.get("consec_capped", True)
            st = r.get("s30_state", "none")
            if st == "none":
                s30_txt = "无记录"
            elif st == "short":
                s30_txt = f"已记录{r.get('s30', 0)}天"
            else:
                s30_txt = r.get("s30", 0)
            if capped:
                days_txt = f"≥{k}天"
                act = f"连续缺货≥{k}天，优先补货/催供"
                why = f"{g}类·连续缺货≥{k}天"
            elif k == 1:
                days_txt = 1
                act = "当天缺货，优先补货/催供"
                why = f"{g}类·当天售罄"
            else:
                days_txt = k
                act = f"连续缺货{k}天，优先补货/催供"
                why = f"{g}类·连续缺货{k}天"
            out.append({"优先级": PRI[g], "等级": g, "商城商品码": str(r["item_id"]), "商品名称": r["name"],
                        "供商": r["provider_name"] or "没有供商", "大分类": r["flow"] + "品",
                        "连续缺货天": days_txt,
                        "售罄天数": s30_txt,
                        "建议动作": act,
                        "入选说明": why})
        return out

    def sheet_data(sname, title, subtitle, data):
        ws = wb.create_sheet(sname)
        title_row(ws, title, subtitle, 10)
        hdr_row(ws, cols10)
        for i, rec in enumerate(data):
            r = 5 + i
            grade_style(ws, r, rec["等级"])
            for j in range(3, 11):
                key = cols10[j - 1]
                c = ws.cell(r, j, rec[key]); c.font = yh(10)
                c.alignment = Alignment(vertical="center", wrap_text=(j in (4, 9, 10)))
        for j, w in enumerate(widths10, 1):
            ws.column_dimensions[chr(64 + j)].width = w
        paint_grid(ws, 4, max(4, 4 + len(data)), 1, 10)
        ws.auto_filter.ref = f"A4:J{max(4, 4 + len(data))}"

    sheet_data("今日必盯", "今日必盯（A）· 建议每天只看这页",
               f"共 {n['A']} 条｜口径 v1.4，基线第 1 天仅 A 可出，B/C/D 待库存历史｜筛选已开启",
               data_rows(["A"]))
    sheet_data("全量名单", "重点关注 · 全量名单（含 C/D）",
               f"共 {len(rows_out)} 条｜按等级 A→D，同级按严重程度降序｜筛选已开启",
               data_rows(None))
    wb.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True, help="ecommerce_bot.yaml 路径")
    ap.add_argument("--date", default=None, help="目标日 YYYY-MM-DD（缺省=MAX(snapshot_date)）")
    ap.add_argument("--outdir", default="./report_out", help="输出目录")
    args = ap.parse_args()

    cfg = load_cfg(args.config)
    tables = cfg["tables"]
    so = cfg["caliber"]["soldout"]
    flow_cats = so["flow_cats"]
    grade_th = {"流通": so["grade"]["flow"], "非流通": so["grade"]["non_flow"]}

    db = DB(cfg)
    db.connect()
    t0 = time.time()
    try:
        # 目标日
        if args.date:
            day = args.date
        else:
            _, r = db.q("order", f"SELECT MAX(snapshot_date) FROM {tables['inventory_sale']} WHERE snapshot_date IS NOT NULL")
            if not r or not r[0][0]:
                raise SystemExit("[FATAL] inventory_sale 无任何快照，终止")
            day = str(r[0][0])
        year_start = day.split("-")[0] + "-01-01"
        year_since = int(time.mktime(time.strptime(year_start, "%Y-%m-%d")))

        # 1) 上架商品
        _, items = db.q("goods", f"SELECT item_id, name, provider_id FROM {tables['item_sale']} WHERE status={so['on_sale_status']} AND sale_type={so['sale_type']}")
        item_map = {int(r[0]): {"name": r[1], "provider_id": int(r[2] or 0)} for r in items}
        print(f"[1/7] 上架商品(status={so['on_sale_status']}&sale_type={so['sale_type']}) {len(item_map)} 个")

        # 2) 供商名称
        _, prov = db.q("goods", f"SELECT provider_id, provider_name FROM {tables['item_provider']} WHERE status=1")
        prov_name = {}
        for r in prov:
            # 云仓后缀在库里存在混写：全角（云仓）、半角(云仓）、(云仓) 都有，必须用正则
            nm = re.sub(r"[（(]\s*云仓\s*[)）]", "", str(r[1] or "")).strip()
            prov_name[str(r[0])] = nm
        print(f"[2/7] 供商名称 {len(prov_name)} 个")

        # 3) 类目
        _, g = db.q("goods", f"SELECT item_id, top_cat FROM {tables['goods']}")
        topcat = {int(r[0]): r[1] for r in g}
        ids = set(topcat.values())
        ph = ",".join(["'%s'" % str(x) for x in ids if x])
        _, cats = db.q("goods", f"SELECT id, name FROM {tables['item_category']} WHERE id IN ({ph})")
        cat_name = {r[0]: r[1] for r in cats}
        print(f"[3/7] 类目映射 {len(cat_name)} 个")

        # 4) 当日快照前置校验
        _, chk = db.q("order", f"SELECT COUNT(*) FROM {tables['inventory_sale']} WHERE snapshot_date='{day}'")
        if not chk or chk[0][0] == 0:
            raise SystemExit(f"[FATAL] {tables['inventory_sale']} 无 {day} 快照，终止")

        # 5) 快照 + 售罄候选（上架供商行）
        _, sale = db.q("order", f"SELECT item_id, provider_id, sellable_qty, warehouse_qty, total_qty, is_sold_out "
                                f"FROM {tables['inventory_sale']} WHERE snapshot_date='{day}'")
        sale_map = {(int(r[0]), int(r[1])): {"total": float(r[4] or 0), "sold": int(r[5] or 0)} for r in sale}
        print(f"[4/7] inventory_sale {day} {len(sale_map)} 行")

        # 售罄判据（v1.5）：上架供商行 is_sold_out=1；上架供商=0 且快照里没有供商 0 行
        # 时算售罄、供商留空，不得改记成快照里的旧供商
        soldout = []
        n_no_zero = 0
        for it, meta in item_map.items():
            pid = meta["provider_id"]
            row = sale_map.get((it, pid))
            hit = row is not None and row["sold"]
            if not hit and pid == 0 and row is None:
                hit = True; n_no_zero += 1
            if hit:
                c = topcat.get(it)
                fname = cat_name.get(str(c)) if c else "?"
                flow = "流通" if fname in flow_cats else "非流通"
                soldout.append({"item_id": it, "name": meta["name"], "provider_id": pid,
                                "provider_name": ("" if pid == 0 else prov_name.get(str(pid), str(pid))),
                                "fname": fname, "flow": flow,
                                "total_inv": (0.0 if row is None else row["total"])})
        print(f"[5/7] 当天售罄（上架供商 is_sold_out=1） {len(soldout)} 个｜其中无供商0行补判 {n_no_zero} 个（供商留空）")

        # 5.5) 连续缺货天回溯（v1.5）：判据与当日售罄完全一致，从目标日往前逐日数到「有货」为止
        _, sd = db.q("order", "SELECT DISTINCT snapshot_date FROM "
                              f"{tables['inventory_sale']} WHERE snapshot_date<='{day}'")
        snap_days = sorted({str(r[0]) for r in sd})
        hist, seen_item = {}, set()
        ids_all = [s["item_id"] for s in soldout]
        for i in range(0, len(ids_all), 400):
            chunk = ids_all[i:i+400]
            ph = ",".join(["'%s'" % x for x in chunk])
            _, hr = db.q("order", "SELECT snapshot_date, item_id, provider_id, is_sold_out "
                                  f"FROM {tables['inventory_sale']} "
                                  f"WHERE snapshot_date BETWEEN '{snap_days[0]}' AND '{day}' "
                                  f"AND item_id IN ({ph})")
            for r in hr:
                hist[(str(r[0]), int(r[1]), int(r[2]))] = int(r[3] or 0)
                seen_item.add((str(r[0]), int(r[1])))

        def soldout_at(it, pid, d):
            if (d, it, pid) in hist:
                return hist[(d, it, pid)] == 1
            return pid == 0          # 无供商 0 行 → 与当日判据一致，算售罄

        d30 = (datetime.date.fromisoformat(day) - datetime.timedelta(days=29)).isoformat()
        win30 = [d for d in snap_days if d >= d30 and d <= day]
        n_capped = n_none = n_short = 0
        for s in soldout:
            it, pid = s["item_id"], s["provider_id"]
            k = 0
            for d in reversed(snap_days):
                if soldout_at(it, pid, d):
                    k += 1
                else:
                    break
            s["consec_days"] = k
            s["consec_capped"] = (k >= len(snap_days))
            if s["consec_capped"]:
                n_capped += 1
            if not win30 or not any((d, it) in seen_item for d in win30):
                s["s30"], s["s30_state"] = 0, "none"
                n_none += 1
            else:
                s["s30"] = sum(1 for d in win30 if soldout_at(it, pid, d))
                if len(win30) >= 30:
                    s["s30_state"] = "ok"
                else:
                    s["s30_state"] = "short"
                    n_short += 1
        print(f"[5.5/7] 连续缺货回溯：快照 {len(snap_days)} 天（{snap_days[0]}~{snap_days[-1]}）"
              f"｜数到边界仍在断 {n_capped} 个（标 ≥N天）"
              f"｜售罄天数窗口 {len(win30)} 天（{win30[0]}~{win30[-1]}）：可给结论 {len(soldout)-n_none} 个、"
              f"窗口未满30天 {n_short} 个（标已记录N天）、无记录 {n_none} 个")

        # 6) 月销（order_code 关联、Unix 时间戳、valid=1）
        monthly = {}
        ids = [s["item_id"] for s in soldout]
        for i in range(0, len(ids), 400):
            chunk = ids[i:i + 400]
            ph = ",".join(["'%s'" % x for x in chunk])
            sql = (f"SELECT od.sku_code, DATE_FORMAT(FROM_UNIXTIME(oh.order_time),'%Y-%m') ym, SUM(od.price*od.quantity) "
                   f"FROM {tables['order_detail']} od JOIN {tables['order_head']} oh ON od.order_code=oh.order_code "
                   f"WHERE od.sku_code IN ({ph}) AND oh.order_time >= {year_since} AND oh.valid=1 "
                   f"GROUP BY od.sku_code, ym")
            _, rows = db.q("order", sql)
            for r in rows:
                monthly.setdefault(r[0], {})[r[1]] = float(r[2])
        print(f"[6/7] 月销聚合完成 {len(monthly)} 个有销售")

        # 7) 组装
        rows_out = []
        for s in soldout:
            ms = monthly.get(s["item_id"], {})
            max_amt = max(ms.values()) if ms else 0.0
            gr = grade_of(max_amt, s["flow"], grade_th)
            rows_out.append({**s, "max_month_amt": max_amt, "max_month": max(ms, key=ms.get) if ms else "-", "grade": gr})
        gorder = {"A": 0, "B": 1, "C": 2, "D": 3}
        rows_out.sort(key=lambda x: (gorder[x["grade"]], -x["max_month_amt"]))
        n = {"A": 0, "B": 0, "C": 0, "D": 0}
        for r in rows_out:
            n[r["grade"]] += 1

        os.makedirs(args.outdir, exist_ok=True)
        dt = {"date": day, "caliber": "v1.4", "soldout_total": len(soldout), "grade_dist": n,
              "elapsed_s": round(time.time() - t0, 2)}
        result = {"meta": dt, "rows": rows_out}
        json_path = os.path.join(args.outdir, f"qpwy_卖断货_{day.replace('-', '')}.json")
        json.dump(result, open(json_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        xlsx_path = os.path.join(args.outdir, f"03-卖断货-重点关注商品_当日快照_{day.replace('-', '')}.xlsx")
        build_xlsx(rows_out, n, day, xlsx_path)

        print(f"[7/7] 完成：售罄 {len(rows_out)}（A{n['A']}/B{n['B']}/C{n['C']}/D{n['D']}）耗时 {dt['elapsed_s']}s")
        print("JSON:", json_path)
        print("XLSX:", xlsx_path)
    finally:
        db.close()


if __name__ == "__main__":
    main()
