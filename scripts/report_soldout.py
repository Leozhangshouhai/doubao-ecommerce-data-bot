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

# 供商列统一文案：没有供商（provider_id=0/空）的品一律标记为「自营」（交付方封板 2026-11-19）
SELF_OP_NAME = "自营"


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


def _hx(c):
    """6 位色值 → xlsxwriter 需要的 #RRGGBB。"""
    return c if c.startswith("#") else "#" + c


# ---------- 第 4 页：全部供应商售罄跟踪（不是第 5 张报表） ----------
TRACK_COLS = ["商城商品码", "商品名称", "一级分类", "大分类", "供商",
              "最高月销", "等级", "近30天售罄天数"]
TRACK_FILL = {"售罄": "FCE4E4", "有货": "E2EFDA", "下架": "EDEDED"}
TRACK_CELL = ["下架", "有货", "售罄"]          # 与矩阵字节值 0/1/2 对应

_LAST_LOG = [time.time()]


def tlog(msg):
    _LAST_LOG[0] = time.time()
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def _watchdog(limit=600, interval=20):
    """任意一步超过 limit 秒没有新日志 → 打印栈并退出，不干等。"""
    def run():
        import traceback
        while True:
            time.sleep(interval)
            if time.time() - _LAST_LOG[0] > limit:
                print(f"[WATCHDOG] {limit}s 无新日志，主动停止。栈：", flush=True)
                for fr in sys._current_frames().values():
                    print("".join(traceback.format_stack(fr)[-8:]), flush=True)
                os._exit(3)
    threading.Thread(target=run, daemon=True).start()


MONTHLY_CACHE = os.path.expanduser("~/.qpwy_bot/soldout_monthly_cache.json")


def build_monthly(db, tables, year_since, day):
    """最高月销：历史月份走缓存，每天只聚合「未入缓存的月份」（通常只有当月）。
    缓存失效（跨年起点变了）才全量重算一次。"""
    cur_ym = day[:7]
    c = None
    if os.path.exists(MONTHLY_CACHE):
        try:
            c = json.load(open(MONTHLY_CACHE, encoding="utf-8"))
            if str(c.get("year_since")) != str(year_since):
                c = None
        except Exception:
            c = None

    if c and c.get("last_ym") and c.get("items"):
        items = c["items"]
        last = c["last_ym"]
        start_ym = last if last < cur_ym else cur_ym
        t0 = time.time()
        ts = int(time.mktime(time.strptime(start_ym + "-01", "%Y-%m-%d")))
        sql = (f"SELECT od.sku_code, DATE_FORMAT(FROM_UNIXTIME(oh.order_time),'%Y-%m') ym, "
               f"SUM(od.price*od.quantity) FROM {tables['order_detail']} od "
               f"JOIN {tables['order_head']} oh ON od.order_code=oh.order_code "
               f"WHERE od.create_time >= {ts} AND oh.order_time >= {year_since} AND oh.valid=1 "
               f"GROUP BY od.sku_code, ym")
        _, rows = db.q("order", sql)
        for r in rows:
            items.setdefault(str(r[0]), {})[r[1]] = float(r[2])
        c["last_ym"] = cur_ym
        c["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")
        json.dump(c, open(MONTHLY_CACHE, "w", encoding="utf-8"), ensure_ascii=False)
        tlog(f"[月销] 增量 {start_ym}~{cur_ym}：聚合 {len(rows)} 行，耗时 {time.time()-t0:.1f}s")
        return items

    t0 = time.time()
    sql = (f"SELECT od.sku_code, DATE_FORMAT(FROM_UNIXTIME(oh.order_time),'%Y-%m') ym, "
           f"SUM(od.price*od.quantity) FROM {tables['order_detail']} od "
           f"JOIN {tables['order_head']} oh ON od.order_code=oh.order_code "
           f"WHERE od.create_time >= {year_since} AND oh.order_time >= {year_since} AND oh.valid=1 "
           f"GROUP BY od.sku_code, ym")
    _, rows = db.q("order", sql)
    items = {}
    for r in rows:
        items.setdefault(str(r[0]), {})[r[1]] = float(r[2])
    os.makedirs(os.path.dirname(MONTHLY_CACHE), exist_ok=True)
    json.dump({"year_since": str(year_since), "last_ym": cur_ym, "items": items,
               "updated": time.strftime("%Y-%m-%d %H:%M:%S")},
              open(MONTHLY_CACHE, "w", encoding="utf-8"), ensure_ascii=False)
    tlog(f"[月销] 全量重建：聚合 {len(rows)} 行，{len(items)} 个商品，耗时 {time.time()-t0:.1f}s")
    return items


def build_track_data(db, tables, so, prov_name, cat_name, topcat,
                     flow_cats, day, year_since, item_map, grade_th, monthly=None):
    """
    行 = 窗口内 hist 出现过的商品 ∪ 目标日仍上架（下架过的也留）。
    列 = 有记录的库存日（新→旧），整天没记录的日子不建列。

    增量：hist 段整段只扫一次（落 soldout_track_hist.json）；快照段落增量状态
    soldout_track_state.json，每天只查「状态里没有的新日期」，历史列直接复用。
    目标日回退或状态版本不匹配时才重建。
    """
    dbname = tables["inventory_sale"].split(".")[0]
    HIST = f"{dbname}.inventory_daily_sale_hist"
    cache_path = os.path.expanduser("~/.qpwy_bot/soldout_track_hist.json")
    state_path = os.path.expanduser("~/.qpwy_bot/soldout_track_state.json")

    d0 = datetime.date.fromisoformat(day)
    yy, mm = d0.year, d0.month - 6
    if mm <= 0:
        yy -= 1; mm += 12
    win_start = datetime.date(yy, mm, min(d0.day, 28)).isoformat()
    if win_start < "2026-07-01":
        win_start = "2026-07-01"

    # ---------- [1] hist 缓存 ----------
    hist_dates, hist_map, hist_last_pid = [], {}, {}
    used_cache = False
    _, hm = db.q("order", f"SELECT MAX(snapshot_date) FROM {HIST}")
    cur_hist_end = str(hm[0][0]) if hm and hm[0][0] is not None else None
    if os.path.exists(cache_path) and cur_hist_end:
        try:
            c = json.load(open(cache_path, encoding="utf-8"))
            if c.get("hist_end") == cur_hist_end:
                hist_dates = c["dates"]
                hist_map = {int(k): v for k, v in c["items"].items()}
                hist_last_pid = {int(k): int(v) for k, v in c.get("last_pid", {}).items()}
                used_cache = True
                tlog(f"[1] 命中 hist 缓存：{len(hist_dates)} 天 / {len(hist_map)} 商品"
                     f"（{hist_dates[0]}~{hist_dates[-1]}），不扫表")
            else:
                tlog(f"[1] hist 表已到 {cur_hist_end}（缓存 {c.get('hist_end')}），重建缓存")
        except Exception as e:
            tlog(f"[1] hist 缓存读取失败，重建：{e}")

    hist_end = hist_dates[-1] if hist_dates else None
    if not used_cache:
        t1 = time.time()
        _, hd = db.q("order", f"SELECT DISTINCT snapshot_date FROM {HIST} "
                              f"WHERE snapshot_date BETWEEN '2026-07-01' AND '{cur_hist_end}'")
        hist_dates = sorted(str(r[0]) for r in hd)
        hist_end = hist_dates[-1] if hist_dates else None
        tlog(f"[1] 建 hist 缓存：库存日 {len(hist_dates)} 天"
             f"（{hist_dates[0] if hist_dates else '-'}~{hist_end or '-'}）")
        if hist_end:
            _, hr = db.q("order", f"SELECT snapshot_date, item_id, is_sold_out FROM {HIST} "
                                  f"WHERE snapshot_date BETWEEN '2026-07-01' AND '{hist_end}'")
            pos = {d: i for i, d in enumerate(hist_dates)}
            buf = {}
            for r in hr:
                it = int(r[1])
                s = buf.get(it)
                if s is None:
                    s = ["0"] * len(hist_dates)
                    buf[it] = s
                s[pos[str(r[0])]] = "2" if int(r[2] or 0) == 1 else "1"
            hist_map = {k: "".join(v) for k, v in buf.items()}
            _, hp = db.q("order", f"SELECT item_id, provider_id FROM {HIST} "
                                  f"WHERE snapshot_date='{hist_end}'")
            hist_last_pid = {int(r[0]): int(r[1] or 0) for r in hp}
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        json.dump({"win_start": win_start, "hist_end": hist_end, "dates": hist_dates,
                   "items": {str(k): v for k, v in hist_map.items()},
                   "last_pid": {str(k): v for k, v in hist_last_pid.items()},
                   "built_at": time.strftime("%Y-%m-%d %H:%M:%S")},
                  open(cache_path, "w", encoding="utf-8"), ensure_ascii=False)
        tlog(f"[1] hist 缓存已落盘，耗时 {time.time()-t1:.1f}s")

    # ---------- [2] 快照段：增量判断 ----------
    t2 = time.time()
    if hist_end:
        cut = (datetime.date.fromisoformat(hist_end) + datetime.timedelta(days=1)).isoformat()
    else:
        cut = "2026-11-03"
    st = None
    if os.path.exists(state_path):
        try:
            s = json.load(open(state_path, encoding="utf-8"))
            if (s.get("hist_end") == hist_end and s.get("dates")
                    and str(s["dates"][-1]) <= day):
                st = s
        except Exception as e:
            tlog(f"[2] 增量状态读取失败，重建：{e}")

    if st and st["dates"][-1] == day:
        base_dates = st["dates"]
        base_items = {int(k): v for k, v in st["items"].items()}
        new_dates = []
        tlog(f"[2] 状态已到目标日 {day}，跳过库存日查询")
    elif st:
        _, sd = db.q("order", f"SELECT DISTINCT snapshot_date FROM {tables['inventory_sale']} "
                              f"WHERE snapshot_date >= '{cut}' AND snapshot_date <= '{day}'")
        sale_dates = sorted(str(r[0]) for r in sd)
        base_dates = st["dates"]
        base_items = {int(k): v for k, v in st["items"].items()}
        new_dates = [d for d in sale_dates if d not in base_dates]
        tlog(f"[2] 增量：已有 {len(base_dates)} 个日列（至 {base_dates[-1]}）"
             f"｜待补 {len(new_dates)} 天 {new_dates if new_dates else ''}，"
             f"耗时 {time.time()-t2:.1f}s")
    else:
        _, sd = db.q("order", f"SELECT DISTINCT snapshot_date FROM {tables['inventory_sale']} "
                              f"WHERE snapshot_date >= '{cut}' AND snapshot_date <= '{day}'")
        sale_dates = sorted(str(r[0]) for r in sd)
        base_dates = list(hist_dates)
        base_items = dict(hist_map)
        new_dates = list(sale_dates)
        tlog(f"[2] 首次建状态：hist {len(hist_dates)} 天 + 快照 {len(new_dates)} 天，"
             f"耗时 {time.time()-t2:.1f}s")

    all_dates = base_dates + new_dates
    if not all_dates:
        tlog("[2] 库存日为空，跳过第 4 页")
        return None
    old_len = len(base_dates)

    # ---------- [3] 行集 + A-H ----------
    t3 = time.time()
    _, cs = db.q("goods", f"SELECT item_id, name, provider_id, status FROM {tables['item_sale']} "
                          f"WHERE status IN ({so['on_sale_status']}, 3) AND sale_type={so['sale_type']}")
    cur_map = {int(r[0]): {"name": r[1], "provider_id": int(r[2] or 0), "st": int(r[3] or 0)}
               for r in cs}
    on_ids = {it for it, v in cur_map.items() if v["st"] == int(so["on_sale_status"])}
    rows_items = sorted(set(hist_map) | set(base_items) | on_ids)
    new_items = [it for it in rows_items if it not in base_items]
    tlog(f"[3] 行集 {len(rows_items)} 个商品｜状态里没有的新商品 {len(new_items)} 个")

    pid_set = sorted({v["provider_id"] for v in cur_map.values()} | set(hist_last_pid.values()))
    ph = ",".join(str(p) for p in pid_set)

    # ① 只查新增日期（通常 1 天）
    filled = {}
    if new_dates:
        qd = ",".join("'%s'" % d for d in new_dates)
        _, sr = db.q("order", f"SELECT snapshot_date, item_id, provider_id, is_sold_out "
                              f"FROM {tables['inventory_sale']} "
                              f"WHERE snapshot_date IN ({qd}) AND provider_id IN ({ph})")
        nidx = {d: i for i, d in enumerate(new_dates)}
        for r in sr:
            it = int(r[1]); pid = cur_map.get(it, {}).get("provider_id", hist_last_pid.get(it, 0))
            if int(r[2] or 0) != pid:
                continue
            filled.setdefault(it, {})[nidx[str(r[0])]] = 2 if int(r[3] or 0) == 1 else 1
        tlog(f"[3] 新增日快照行 {len(sr)}｜命中 {len(filled)} 个商品")

    # ② 新商品补历史快照列（只补它自己那几行，量很小）
    if new_items and old_len > len(hist_dates):
        ni = ",".join(str(x) for x in new_items)
        old_sale = [d for d in base_dates[len(hist_dates):]]
        _, nr = db.q("order", f"SELECT snapshot_date, item_id, provider_id, is_sold_out "
                              f"FROM {tables['inventory_sale']} "
                              f"WHERE snapshot_date >= '{cut}' AND snapshot_date <= '{day}' "
                              f"AND item_id IN ({ni}) AND provider_id IN ({ph})")
        oi = {d: i for i, d in enumerate(old_sale)}
        for r in nr:
            it = int(r[1]); pid = cur_map.get(it, {}).get("provider_id", hist_last_pid.get(it, 0))
            if int(r[2] or 0) != pid:
                continue
            d = str(r[0])
            if d in oi:
                filled.setdefault(it, {})["H%d" % oi[d]] = 2 if int(r[3] or 0) == 1 else 1
        tlog(f"[3] 新商品补历史 {len(nr)} 行")

    # ③ 拼矩阵（历史列直接复用，只动新列）
    items_s = {}
    nd = len(new_dates)
    for it in rows_items:
        base = base_items.get(it)
        if base is None:
            h = hist_map.get(it)
            if h is not None:
                base = h + "0" * (old_len - len(h))
            else:
                base = "0" * old_len
        fx = filled.get(it, {})
        add = []
        for j in range(nd):
            v = fx.get(j)
            if v is None:
                v = fx.get("H%d" % (old_len + j)) if False else None
            add.append("0" if v is None else ("2" if v == 2 else "1"))
        items_s[it] = base + "".join(add)

    # ④ 新商品的历史旧快照列补进 base（上面按 H 前缀暂存）
    if new_items and old_len > len(hist_dates):
        for it in new_items:
            fx = filled.get(it, {})
            pref = {k: v for k, v in fx.items() if isinstance(k, str)}
            if not pref:
                continue
            cur = items_s[it]
            lst = list(cur)
            for k, v in pref.items():
                lst[len(hist_dates) + int(k[1:])] = "2" if v == 2 else "1"
            items_s[it] = "".join(lst)

    json.dump({"win_start": win_start, "hist_end": hist_end, "dates": all_dates,
               "items": {str(k): v for k, v in items_s.items()},
               "built_at": time.strftime("%Y-%m-%d %H:%M:%S")},
              open(state_path, "w", encoding="utf-8"), ensure_ascii=False)
    tlog(f"[3] 增量状态已落盘（{len(all_dates)} 列 / {len(items_s)} 商品）")

    # ⑤ 窗口左边界随时间前移时，只裁输出列，缓存仍保留全量（不触发重建）
    if all_dates and all_dates[0] < win_start:
        keep = [i for i, d in enumerate(all_dates) if d >= win_start]
        for it in list(items_s):
            items_s[it] = "".join(items_s[it][i] for i in keep)
        all_dates = [all_dates[i] for i in keep]
        tlog(f"[2] 窗口从 {win_start} 起，输出裁到 {len(all_dates)} 列（缓存不重建）")

    # ⑥ 近 30 天 / 月销 / A-H
    n30 = min(30, len(all_dates))
    if monthly is None:
        monthly = build_monthly(db, tables, year_since, day)
        tlog(f"[3] 月销（缓存）{len(monthly)} 个商品")
    else:
        tlog(f"[3] 月销复用前三页结果：{len(monthly)} 个商品，0s")

    goods_name = {it: v.get("name") for it, v in item_map.items() if v.get("name")}

    track_rows = []
    for it in rows_items:
        s = items_s[it]
        pid = cur_map.get(it, {}).get("provider_id", hist_last_pid.get(it, 0))
        nm = cur_map.get(it, {}).get("name") or goods_name.get(it) or ""
        c = topcat.get(it)
        fname = cat_name.get(str(c)) if c else "?"
        flow = "流通" if fname in flow_cats else "非流通"
        ms = monthly.get(str(it), {})
        amt = max(ms.values()) if ms else 0.0
        s30 = s[-n30:].count("2")
        track_rows.append({"item_id": str(it), "name": nm, "fname": fname or "?",
                           "flow": flow,
                           "provider_name": (SELF_OP_NAME if pid == 0 else (prov_name.get(str(pid)) or SELF_OP_NAME)),
                           "max_amt": round(amt, 2), "grade": grade_of(amt, flow, grade_th),
                           "s30": s30, "s": s})
    go = {"A": 0, "B": 1, "C": 2, "D": 3}
    track_rows.sort(key=lambda x: (go[x["grade"]], -x["max_amt"]))
    tlog(f"[3] A-H 拼装完成 {len(track_rows)} 行，耗时 {time.time()-t3:.1f}s")
    return {"dates": all_dates, "rows": track_rows}


def build_xlsx(rows_out, n, day, out_path, track=None):
    """用 xlsxwriter 出表：前三页逐格保真，第四页 197 万格一次 write_row 写入。

    openpyxl 写 197 万格 + 保存约 100s，xlsxwriter 约 30s，文件也从 6.8MB 降到 ~4MB。
    """
    import xlsxwriter
    t_x = time.time()
    book = xlsxwriter.Workbook(out_path, {"strings_to_numbers": False,
                                          "strings_to_formulas": False,
                                          "strings_to_urls": False})

    F = {}

    def fmt(key, **kw):
        if key in F:
            return F[key]
        base = {"font_name": "Microsoft YaHei", "font_size": 10, "valign": "vcenter"}
        base.update(kw)
        F[key] = book.add_format(base)
        return F[key]

    # 通用格式
    fmt("title", font_size=16, bold=True, color=_hx(C["title"]))
    fmt("sub", font_size=10, color=_hx(C["subtitle"]))
    fmt("hdr", font_size=11, bold=True, color="#FFFFFF", bg_color=_hx(C["hdr"]),
        align="center", border=1, border_color="#E5E7EB")
    fmt("body", border=1, border_color="#E5E7EB")
    fmt("body_wrap", border=1, border_color="#E5E7EB", text_wrap=True)
    fmt("num", border=1, border_color="#E5E7EB", align="right", num_format="#,##0")
    fmt("center", border=1, border_color="#E5E7EB", align="center")
    fmt("plain", )
    fmt("plain_b", bold=True)
    fmt("plain_c", align="center")
    fmt("plain_cb", bold=True, align="center")
    fmt("right", num_format="#,##0", align="right")
    for g in "ABCD":
        fmt(f"gpri_{g}", bold=True, color="#FFFFFF", bg_color=_hx(deep(g)),
            align="center", border=1, border_color="#E5E7EB")
        fmt(f"g_{g}", bold=True, color=_hx(tone(g)), bg_color=_hx(light(g)),
            align="center", border=1, border_color="#E5E7EB")
        fmt(f"gnum_{g}", font_size=22, bold=True, color=_hx(tone(g)),
            bg_color=_hx(light(g)), align="center", border=1, border_color="#E5E7EB")
        fmt(f"glab_{g}", bold=True, color=_hx(tone(g)), align="center")
        fmt(f"gdes_{g}", color=_hx(C["subtitle"]), align="center")
    # 第四页
    fmt("t_hdr", font_size=10, bold=True, color="#FFFFFF", bg_color=_hx(C["hdr"]),
        align="center", border=1, border_color="#E5E7EB")
    fmt("t_datehdr", font_size=9, bold=True, color="#FFFFFF", bg_color=_hx(C["hdr"]),
        align="center")
    fmt("t_body", border=1, border_color="#E5E7EB")
    fmt("t_num", border=1, border_color="#E5E7EB", align="right", num_format="#,##0")
    fmt("t_center", border=1, border_color="#E5E7EB", align="center")
    fmt("t_date", font_size=9, align="center", border=1, border_color="#E5E7EB")

    # ---------- 页 1：一眼看懂 ----------
    ws = book.add_worksheet("一眼看懂")
    ws.set_row(0, 28)
    ws.merge_range(0, 0, 0, 5, "重点关注商品 · 当日快照", F["title"])
    ws.merge_range(1, 0, 1, 5,
                   f"数据日期 {day}｜共 {len(rows_out)} 个 SKU｜口径 v1.4（inventory_daily_sale 上架供商）",
                   F["sub"])
    ws.write(3, 0, "按等级数量", F["plain_b"])
    gname = {"A": "今日必盯 A", "B": "持续缺货 B", "C": "频发缺货 C", "D": "低销频缺 D"}
    gdes = {"A": "月销>1万", "B": "月销1000~1万", "C": "月销500~1000", "D": "月销<500"}
    ws.set_row(4, 40)
    for i, g in enumerate("ABCD"):
        ws.write_number(4, i, n[g], F[f"gnum_{g}"])
        ws.write(5, i, gname[g], F[f"glab_{g}"])
        ws.write(6, i, gdes[g], F[f"gdes_{g}"])
    ws.write(8, 0, "怎么用（建议）", F["plain_b"])
    notes = [
        "1. 等级按月销额分档：A 流通>1万/非流通>2千，B 1000+、C 500+、D<500；与缺货天数无关。",
        "2. 「全量名单」按优先级排序，都是当日售罄的品；颜色：红=A 橙=B 黄=C 灰=D。",
        "3. 「连续缺货天」与「售罄天数」都按已有库存记录逐日实算；数到最早一天仍在断的标「≥N天」，显示「已记录N天」表示还没攒满 30 天、暂不给结论。",
        "4. D 类默认放全量表末尾；日常可不展开，需要时再筛供商。",
        "5. 「建议动作」给操作指引，规则细节仍保留在「入选说明」列。",
    ]
    for i, nt in enumerate(notes):
        ws.merge_range(9 + i, 0, 9 + i, 5, nt, F["plain"])
    ws.write(15, 0, "今日必盯供商 TOP（A+B）", F["plain_b"])
    for j, h in enumerate(["供商", "A+B 数量", "其中 A", "其中 B"]):
        ws.write(16, j, h, F["hdr"])
    ab = {}
    for r in rows_out:
        if r["grade"] in ("A", "B"):
            p = r["provider_name"] or SELF_OP_NAME
            ab.setdefault(p, [0, 0, 0])
            ab[p][0] += 1
            ab[p][1 if r["grade"] == "A" else 2] += 1
    top = sorted(ab.items(), key=lambda x: (-x[1][0], -x[1][1]))[:10]
    for i, (pname, (tot, na, nb)) in enumerate(top):
        r = 17 + i
        ws.write(r, 0, pname, F["body"])
        ws.write_number(r, 1, tot, F["center"])
        ws.write_number(r, 2, na, F["center"])
        ws.write_number(r, 3, nb, F["center"])
    for j in range(4):
        ws.set_column(j, j, 22)

    cols10 = ["优先级", "等级", "商城商品码", "商品名称", "供商", "大分类", "连续缺货天", "售罄天数", "建议动作", "入选说明"]
    widths10 = [12, 6, 12, 36, 28, 10, 15, 18, 26, 28]

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
                        "供商": r["provider_name"] or SELF_OP_NAME, "大分类": r["flow"] + "品",
                        "连续缺货天": days_txt, "售罄天数": s30_txt,
                        "建议动作": act, "入选说明": why})
        return out

    def sheet_data(sname, title, subtitle, data):
        w = book.add_worksheet(sname)
        w.set_row(0, 28)
        w.merge_range(0, 0, 0, 9, title, F["title"])
        w.merge_range(1, 0, 1, 9, subtitle, F["sub"])
        for j, h in enumerate(cols10):
            w.write(3, j, h, F["hdr"])
        for i, rec in enumerate(data):
            r = 4 + i
            g = rec["等级"]
            w.write(r, 0, rec["优先级"], F[f"gpri_{g}"])
            w.write(r, 1, g, F[f"g_{g}"])
            w.write_string(r, 2, rec["商城商品码"], F["body"])
            w.write(r, 3, rec["商品名称"], F["body_wrap"])
            w.write(r, 4, rec["供商"], F["body"])
            w.write(r, 5, rec["大分类"], F["body"])
            if isinstance(rec["连续缺货天"], int):
                w.write_number(r, 6, rec["连续缺货天"], F["center"])
            else:
                w.write(r, 6, rec["连续缺货天"], F["center"])
            if isinstance(rec["售罄天数"], int):
                w.write_number(r, 7, rec["售罄天数"], F["center"])
            else:
                w.write(r, 7, rec["售罄天数"], F["center"])
            w.write(r, 8, rec["建议动作"], F["body_wrap"])
            w.write(r, 9, rec["入选说明"], F["body_wrap"])
        for j, wd in enumerate(widths10):
            w.set_column(j, j, wd)
        last = max(3, 3 + len(data))
        w.autofilter(3, 0, last, 9)

    sheet_data("今日必盯", "今日必盯（A）· 建议每天只看这页",
               f"共 {n['A']} 条｜口径 v1.4，基线第 1 天仅 A 可出，B/C/D 待库存历史｜筛选已开启",
               data_rows(["A"]))
    sheet_data("全量名单", "重点关注 · 全量名单（含 C/D）",
               f"共 {len(rows_out)} 条｜按等级 A→D，同级按严重程度降序｜筛选已开启",
               data_rows(None))

    if track:
        build_track_sheet(book, track, day, F)

    tlog("[6] 开始保存 xlsx（前三页 + 第四页，xlsxwriter）")
    book.close()
    tlog(f"[6] 已保存 {out_path}，出表耗时 {time.time()-t_x:.1f}s")
    return out_path


def build_track_sheet(book, track, day, F):
    """第四页：197 万格按行一次写入；颜色走条件格式，不为每格建样式。"""
    tlog("[4] 开始写第四页")
    ws = book.add_worksheet("全部供应商售罄跟踪")
    dates = track["dates"]
    rev = list(reversed(dates))
    rows = track["rows"]
    nfix = len(TRACK_COLS)
    nd = len(rev)
    lastc = nfix + nd
    di = {d: i for i, d in enumerate(dates)}
    pos = [di[d] for d in rev]

    ws.set_row(0, 28)
    ws.merge_range(0, 0, 0, min(lastc - 1, 23), "全部供应商售罄跟踪", F["title"])
    ws.merge_range(1, 0, 1, min(lastc - 1, 23),
                   f"数据日期 {day}｜共 {len(rows)} 个商品｜{nd} 个库存日"
                   f"（{dates[0]}~{dates[-1]}）｜快照日按当前上架供商匹配，供商中途换过的日子可能显示为下架",
                   F["sub"])

    for j, h in enumerate(TRACK_COLS):
        ws.write(2, j, h, F["t_hdr"])
    for j, d in enumerate(rev):
        ws.write(2, nfix + j, d[5:], F["t_datehdr"])

    f_body = F["t_body"]
    f_num = F["t_num"]
    f_ctr = F["t_center"]
    f_date = F["t_date"]
    cells = TRACK_CELL                       # ["下架","有货","售罄"]
    n = len(rows)
    r0 = 3
    for i, r in enumerate(rows):
        rr = r0 + i
        ws.write_string(rr, 0, r["item_id"], f_body)
        ws.write_string(rr, 1, r["name"] or "", f_body)
        ws.write_string(rr, 2, r["fname"] or "", f_body)
        ws.write_string(rr, 3, r["flow"] + "品", f_body)
        ws.write_string(rr, 4, r["provider_name"] or SELF_OP_NAME, f_body)
        ws.write_number(rr, 5, r["max_amt"], f_num)
        g = r["grade"]
        ws.write_string(rr, 6, g, F[f"g_{g}"])
        ws.write_number(rr, 7, r["s30"], f_ctr)
        s = r["s"]
        vals = [cells[int(s[pos[j]])] for j in range(nd)]
        ws.write_row(rr, nfix, vals, f_date)
        if (i + 1) % 2000 == 0:
            tlog(f"[5] 已写 {i+1}/{n}")
        elif i == 0:
            tlog(f"[5] 已写 1/{n}")

    for j, w in enumerate([13, 32, 14, 11, 26, 13, 7, 15]):
        ws.set_column(j, j, w)
    ws.set_column(nfix, lastc - 1, 6.5)

    ws.freeze_panes(3, 8)
    ws.autofilter(2, 0, r0 + n - 1, 7)

    lc_name = xlsxwriter_utils_col(lastc)
    c1 = xlsxwriter_utils_col(nfix + 1)
    rng = f"{c1}{r0+1}:{lc_name}{r0+n}"
    for label, color in TRACK_FILL.items():
        ws.conditional_format(rng, {"type": "cell", "criteria": "==",
                                    "value": f'"{label}"',
                                    "format": book.add_format({"bg_color": _hx(color)})})

    nsz = sum(r["s"].count("2") for r in rows)
    tlog(f"[6] 第四页写完：{n} 行 × {nd} 列｜售罄格 {nsz}")


def xlsxwriter_utils_col(idx):
    """1 基列号 → 字母（I、DZ …）。"""
    from xlsxwriter.utility import xl_col_to_name
    return xl_col_to_name(idx - 1)


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
    _watchdog(600, 20)
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
        # 时算售罄、供商标「自营」，不得改记成快照里的旧供商
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
                                "provider_name": (SELF_OP_NAME if pid == 0 else (prov_name.get(str(pid)) or SELF_OP_NAME)),
                                "fname": fname, "flow": flow,
                                "total_inv": (0.0 if row is None else row["total"])})
        print(f"[5/7] 当天售罄（上架供商 is_sold_out=1） {len(soldout)} 个｜其中无供商0行补判 {n_no_zero} 个（供商标「自营」）")

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
        monthly = build_monthly(db, tables, year_since, day)
        tlog(f"[6/7] 月销就绪，覆盖 {len(monthly)} 个有销售商品")

        # 7) 组装
        rows_out = []
        for s in soldout:
            ms = monthly.get(str(s["item_id"]), {})
            max_amt = max(ms.values()) if ms else 0.0
            gr = grade_of(max_amt, s["flow"], grade_th)
            rows_out.append({**s, "max_month_amt": max_amt, "max_month": max(ms, key=ms.get) if ms else "-", "grade": gr})
        gorder = {"A": 0, "B": 1, "C": 2, "D": 3}
        rows_out.sort(key=lambda x: (gorder[x["grade"]], -x["max_month_amt"]))
        n = {"A": 0, "B": 0, "C": 0, "D": 0}
        for r in rows_out:
            n[r["grade"]] += 1

        # 8) 第 4 页：全部供应商售罄跟踪（不是第 5 张报表，不改前三页口径）
        tlog("[3] 开始拼第 4 页数据")
        try:
            track = build_track_data(db, tables, so, prov_name, cat_name, topcat,
                                     flow_cats, day, year_since, item_map, grade_th,
                                     monthly=monthly)
        except Exception as e:
            import traceback
            tlog(f"[!] 第 4 页失败，前三页照常交付：{e}")
            traceback.print_exc()
            track = None

        os.makedirs(args.outdir, exist_ok=True)
        dt = {"date": day, "caliber": "v1.4", "soldout_total": len(soldout), "grade_dist": n,
              "elapsed_s": round(time.time() - t0, 2)}
        result = {"meta": dt, "rows": rows_out}
        json_path = os.path.join(args.outdir, f"qpwy_卖断货_{day.replace('-', '')}.json")
        json.dump(result, open(json_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        tlog("[6] JSON 已落盘，开始出 XLSX")
        xlsx_path = os.path.join(args.outdir, f"03-卖断货-重点关注商品_当日快照_{day.replace('-', '')}.xlsx")
        build_xlsx(rows_out, n, day, xlsx_path, track)

        print(f"[7/7] 完成：售罄 {len(rows_out)}（A{n['A']}/B{n['B']}/C{n['C']}/D{n['D']}）耗时 {dt['elapsed_s']}s")
        print("JSON:", json_path)
        print("XLSX:", xlsx_path)
    finally:
        db.close()


if __name__ == "__main__":
    main()
