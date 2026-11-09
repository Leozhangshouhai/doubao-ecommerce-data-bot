#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
爆卖双榜（报表 ②）取数器。只读。

产出 JSON（供 scripts/render_hot_sale.py 出表），不直接出 XLSX。
口径见 sql_templates/hot_sale.yaml 与 references/style-contract.md「② 爆卖」。
**严禁本脚本之外的临时脚本口径**；改口径先改本文件与 yaml，再重跑历史日对比。

用法：
  python3 report_hot_sale.py --config /path/ecommerce_bot.yaml \
      --outdir "../04_样板案例/QPWY/02_报表-爆卖/日报" [--date YYYY-MM-DD]

窗口 = 目标库存日往前 7 个自然日（不含今天，含目标日）。
库存快照不足时照实用已有天数算，series 里没有快照的日子按「有货」处理（不虚报断货）。
"""
import argparse
import json
import os
import socket
import sys
import threading
import time
from datetime import date, datetime, timedelta

ARGS = None


# ------------------------------------------------------------------ 连接
def bridge(lp, rhost, rport, transport):
    def h(c):
        try:
            ch = transport.open_channel("direct-tcpip", (rhost, rport), c.getpeername())
        except Exception:
            c.close()
            return

        def pump(src, dst):
            try:
                while True:
                    d = src.recv(4096)
                    if not d:
                        break
                    dst.send(d)
            except Exception:
                pass
            finally:
                try:
                    dst.close()
                except Exception:
                    pass
                try:
                    src.close()
                except Exception:
                    pass

        threading.Thread(target=pump, args=(c, ch), daemon=True).start()
        pump(ch, c)

    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    srv.bind(("127.0.0.1", lp))
    srv.listen(16)
    while True:
        c, _ = srv.accept()
        threading.Thread(target=h, args=(c,), daemon=True).start()


class DB(object):
    def __init__(self, cfg):
        self.conn = cfg["conn"]
        self.db = cfg["db"]
        self.tunnel_ports = {"order": 13307, "goods": 13306}
        if self.conn.get("tunnel"):
            self.tunnel_ports = {
                "order": self.conn["tunnel"].get("order_port", 13307),
                "goods": self.conn["tunnel"].get("goods_port", 13306),
            }
        self.cli = None
        self.sql_log = []

    def _port(self, kind):
        if self.conn["mode"] == "jump":
            return self.tunnel_ports.get(kind, 13307 if kind == "order" else 13306)
        return int(self.conn["direct"].get(kind + "_port", 3306))

    def _host(self, kind):
        if self.conn["mode"] == "jump":
            return "127.0.0.1"
        return self.conn["direct"].get(kind + "_host", "127.0.0.1")

    def connect(self):
        if self.conn["mode"] != "jump":
            return
        import paramiko
        cli = paramiko.SSHClient()
        cli.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        cli.connect(self.conn["jump_host"], username=self.conn["jump_user"],
                    password=self.conn["jump_password"], timeout=25)
        t = cli.get_transport()
        for kind in ("order", "goods"):
            tg = self.conn["targets"][kind]
            threading.Thread(target=bridge,
                             args=(self._port(kind), tg["host"], tg["port"], t),
                             daemon=True).start()
            time.sleep(0.4)
        self.cli = cli

    def q(self, kind, sql, note=""):
        import pymysql
        self.sql_log.append((kind, note, sql))
        conn = pymysql.connect(host=self._host(kind), port=self._port(kind),
                               user=self.db["user"], password=self.db["password"],
                               connect_timeout=15, charset="utf8mb4")
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        conn.close()
        return rows

    def close(self):
        if self.cli:
            try:
                self.cli.close()
            except Exception:
                pass


# ------------------------------------------------------------------ 指标
def acc_ratio(series):
    """近3日均 ÷ 前4日均。前4日为0且有量 → None（记「新增」）。"""
    pre = series[:4]
    post = series[4:]
    pa = sum(pre) / 4.0
    a3 = sum(post) / 3.0
    if pa <= 0:
        return None if a3 > 0 else 0.0
    return round(a3 / pa, 2)


def rise_days(series, sold_out_by_day=None):
    """上涨天数 = 相邻两天销量递增的次数。断货日销量为 0，天然不贡献上涨。"""
    return sum(1 for i in range(1, len(series)) if series[i] > series[i - 1])


def is_spike(series):
    """单日峰值超过其余 6 日均值 5 倍，且峰值 ≥10 件 → 视为单日爆单尖峰。"""
    s = series
    if max(s) < 10:
        return False
    peak = max(range(7), key=lambda k: s[k])
    rest = [v for k, v in enumerate(s) if k != peak]
    base = sum(rest) / len(rest)
    return base > 0 and s[peak] / base > 5


def grade_of(max_amt, flow, th):
    if max_amt > th[flow]["A"]:
        return "A"
    if max_amt >= th[flow]["B"]:
        return "B"
    if max_amt >= th[flow]["C"]:
        return "C"
    return "D"


SIZE_TIER = [(500, "爆款"), (200, "旺销"), (100, "畅销"), (0, "起量")]
SPEED_TIER = [(10, "黑马"), (5, "猛涨"), (3, "起势")]


def size_tier_of(q):
    for th, name in SIZE_TIER:
        if q >= th:
            return name
    return "起量"


def speed_tier_of(acc):
    if acc is None:
        return "新品"
    for th, name in SPEED_TIER:
        if acc >= th:
            return name
    return "起势"


# ------------------------------------------------------------------ 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True, help="ecommerce_bot.yaml 路径")
    ap.add_argument("--outdir", default="./report_out")
    ap.add_argument("--date", default=None, help="目标库存日 YYYY-MM-DD，默认取最新快照日")
    ap.add_argument("--sql-log", default=None, help="把本次实际执行的 SQL 落盘，默认写到 outdir/logs")
    a = ap.parse_args()

    import yaml
    with open(a.config, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    tables = cfg["tables"]
    so = cfg["caliber"]["soldout"]
    flow_cats = list(so["flow_cats"])
    th = {"流通": so["grade"]["flow"], "非流通": so["grade"]["non_flow"]}

    os.makedirs(a.outdir, exist_ok=True)
    db = DB(cfg)
    t0 = time.time()
    db.connect()
    try:
        # 1) 目标日
        rows = db.q("order", "SELECT MAX(snapshot_date) FROM %s WHERE snapshot_date IS NOT NULL" % tables["inventory_sale"], "目标库存日")
        target = str(rows[0][0])
        if a.date:
            target = a.date
        print("[1/7] 目标日 %s" % target)

        from datetime import datetime as _dt
        tdate = _dt.strptime(target, "%Y-%m-%d").date()
        days = [str(tdate - timedelta(days=i)) for i in range(6, -1, -1)]
        # 库存快照可用日（只会有最近几天）
        hav = db.q("order", "SELECT DISTINCT snapshot_date FROM %s WHERE snapshot_date BETWEEN '%s' AND '%s' ORDER BY snapshot_date" % (
            tables["inventory_sale"], days[0], days[-1]), "窗口内库存快照日")
        snap_days = [str(r[0]) for r in hav]
        print("[2/7] 窗口 %s ~ %s，库存快照可用 %d 天：%s" % (days[0], days[-1], len(snap_days), ",".join(snap_days)))

        # 2) 上架商品
        on = db.q("goods", "SELECT item_id, name, provider_id FROM %s WHERE status=%s AND sale_type=%s" % (
            tables["item_sale"], int(so["on_sale_status"]), int(so["sale_type"])), "上架商品")
        on_shelf = len(on)
        items = {}
        for iid, nm, pid in on:
            it = items.setdefault(int(iid), {"name": nm, "pids": []})
            it["pids"].append(int(pid or 0))
        for it in items.values():
            it["pids"] = sorted(set(it["pids"]))
        for it in items.values():
            it["pid"] = 0 if 0 in it["pids"] else it["pids"][0]
        print("[3/7] 上架商品 %d" % on_shelf)

        # 3) 逐日销量
        sql = """SELECT od.sku_code, DATE(FROM_UNIXTIME(oh.order_time)) AS d, SUM(od.quantity) AS q
FROM {od} od JOIN {oh} oh ON od.order_code = oh.order_code
WHERE oh.valid = 1 AND oh.order_time >= {ts0}
GROUP BY od.sku_code, DATE(FROM_UNIXTIME(oh.order_time))""".format(
            od=tables["order_detail"], oh=tables["order_head"],
            ts0=int(datetime.strptime(days[0], "%Y-%m-%d").replace(tzinfo=None).timestamp()) - 8 * 3600)
        sale_rows = db.q("order", sql, "逐日销量")
        daily = {}
        for code, d, q in sale_rows:
            try:
                iid = int(code)
            except Exception:
                continue
            daily.setdefault(iid, {})[str(d)] = float(q or 0)
        print("[4/7] 有销量商品 %d" % len(daily))

        # 4) 库存快照（窗口内）
        inv = db.q("order", "SELECT snapshot_date, item_id, provider_id, total_qty, is_sold_out FROM %s WHERE snapshot_date BETWEEN '%s' AND '%s'" % (
            tables["inventory_sale"], days[0], days[-1]), "库存快照")
        invmap = {}
        for d, iid, pid, tq, so_flag in inv:
            invmap[(int(iid), int(pid or 0), str(d))] = (float(tq or 0), int(so_flag or 0))
        print("[5/7] 库存快照行 %d" % len(invmap))

        # 5) 月销（自然年，取最高月）
        y0 = date(int(target[:4]), 1, 1)
        ts_y = int(datetime(y0.year, y0.month, y0.day).timestamp()) - 8 * 3600
        sql2 = """SELECT od.sku_code, DATE_FORMAT(FROM_UNIXTIME(oh.order_time), '%%Y-%%m') AS m,
       SUM(od.price * od.quantity) AS amt
FROM {od} od JOIN {oh} oh ON od.order_code = oh.order_code
WHERE oh.valid = 1 AND oh.order_time >= {ts}
GROUP BY od.sku_code, DATE_FORMAT(FROM_UNIXTIME(oh.order_time), '%%Y-%%m')""".format(
            od=tables["order_detail"], oh=tables["order_head"], ts=ts_y)
        amt = {}
        for code, m, v in db.q("order", sql2.replace("%%", "%"), "月销"):
            try:
                iid = int(code)
            except Exception:
                continue
            amt[iid] = max(amt.get(iid, 0.0), float(v or 0))
        print("[6/7] 有月销商品 %d" % len(amt))

        # 6) 分类 + 供商名
        ids = ",".join(str(x) for x in items.keys())
        cats = {}
        topcat = {int(r[0]): r[1] for r in db.q("goods", "SELECT item_id, top_cat FROM %s" % tables["goods"], "goods.top_cat")}
        cids = sorted({str(v) for v in topcat.values() if v})
        cat_name = {}
        for bi in range(0, len(cids), 400):
            chunk = ",".join("'%s'" % x for x in cids[bi:bi + 400])
            for cid, cname in db.q("goods", "SELECT id, name FROM %s WHERE id IN (%s)" % (
                    tables["item_category"], chunk), "类目名"):
                cat_name[cid] = cname or ""
        for iid in items:
            cid = topcat.get(iid)
            if cid is not None:
                cats[iid] = cat_name.get(cid, "")
        pids = sorted({p for v in items.values() for p in v["pids"] if p})
        provs = {}
        if pids:
            for pid, pname in db.q("goods", "SELECT provider_id, provider_name FROM %s WHERE status=1 AND provider_id IN (%s)" % (
                    tables["item_provider"], ",".join(str(x) for x in pids)), "供商名"):
                provs[int(pid)] = pname or ""
        print("[7/7] 分类 %d｜供商 %d" % (len(cats), len(provs)))
    finally:
        db.close()

    # ---------------- 计算
    pool = 0
    sold_out_today = 0
    size_rows, speed_rows, new_rows = [], [], []

    for iid, meta in items.items():
        d = daily.get(iid, {})
        series = [float(d.get(x, 0)) for x in days]
        pid = meta["pid"]
        pids_all = meta["pids"]
        # 逐日有货状态（用于上涨天数/新品档/断货天数）：只认快照明确标记，
        # 快照里没有这一行 → 视为有货（不虚报断货）。任一上架供商行有货即整件未断。
        out_by_day = []
        alive_pid = pid
        for x in days:
            alive = None
            for p in pids_all:
                key = (iid, p, x)
                if key in invmap and invmap[key][1] == 0 and alive is None:
                    alive = (p, invmap[key][0])
            out_by_day.append(0 if alive else (1 if any((iid, p, x) in invmap for p in pids_all) else 0))
            if alive:
                alive_pid = alive[0]
        pid = alive_pid

        # 当日售罄排除：与报表①同一口径（上架供商=0 且快照无供商 0 行也算售罄）
        today = days[-1]
        hit_rows = [p for p in pids_all if (iid, p, today) in invmap]
        if any(invmap[(iid, p, today)][1] == 0 for p in hit_rows):
            today_out = 0
        elif hit_rows:
            today_out = 1
        else:
            today_out = 1 if 0 in pids_all else 0
        q7 = sum(series)
        if today_out == 1:
            sold_out_today += 1
            continue                      # 当日售罄的品不进榜
        if q7 < 10:
            continue
        pool += 1

        a3 = round(sum(series[4:]) / 3.0, 2)
        a4 = sum(series[:4]) / 4.0
        acc = acc_ratio(series)
        rise = rise_days(series, out_by_day)
        out5 = sum(out_by_day[-5:]) if len(out_by_day) >= 5 else sum(out_by_day)
        rec = {
            "item_id": iid, "name": meta["name"], "provider_id": pid,
            "provider_name": provs.get(pid, ""),
            "fname": cats.get(iid, ""),
            "flow": "流通" if cats.get(iid, "") in flow_cats else "非流通",
            "qty_7d": int(q7), "avg_3d": a3, "acc": acc, "rise_days": rise,
            "series": series, "out_days_5": out5,
            "max_month_amt": round(amt.get(iid, 0.0), 2),
        }
        rec["stock_grade"] = grade_of(rec["max_month_amt"], rec["flow"], th)
        rec["cover_days"] = round(invmap.get((iid, pid, days[-1]), (0, 0))[0] / a3, 1) if a3 > 0 else 0
        rec["avg_7d"] = round(q7 / 7.0, 2)

        # 规模榜
        if q7 >= 50 and acc is not None and acc >= 1.5 and rise >= 3:
            r = dict(rec)
            r["tier"] = size_tier_of(q7)
            size_rows.append(r)
        # 速度榜（不设规模上限）
        if (a3 >= 5 and acc is not None and acc >= 3 and rise >= 3
                and not is_spike(series) and all(v > 0 for v in series[4:])):
            r = dict(rec)
            r["tier"] = speed_tier_of(acc)
            speed_rows.append(r)
        # 新品档：前4日为0 + 后3日连续 + 末日不回吐
        valid_idx = [i for i, x in enumerate(days) if x in snap_days]
        early_out = any(out_by_day[i] == 1 for i in valid_idx[:2])
        if (sum(series[:4]) == 0 and a3 >= 5
                and all(v > 0 for v in series[4:])
                and series[6] >= max(series[4], series[5]) * 0.6
                and not early_out):
            r = dict(rec)
            r["tier"] = "新品"
            r["acc"] = None
            new_rows.append(r)

    order = {"爆款": 0, "旺销": 1, "畅销": 2, "起量": 3}
    size_rows.sort(key=lambda x: (order.get(x["tier"], 9), -x["qty_7d"]))
    sorder = {"黑马": 0, "猛涨": 1, "起势": 2}
    speed_rows.sort(key=lambda x: (sorder.get(x["tier"], 9), -(x["acc"] or 0)))
    new_rows.sort(key=lambda x: -x["qty_7d"])

    payload = {
        "meta": {
            "day": target,
            "window": [days[0], days[-1]],
            "source": "dmall_saas_oms.order_detail/order_head + dmall_saas_atp.inventory_daily_sale",
            "caliber": ("规模:近7日≥50且加速≥1.5且涨≥3｜"
                        "速度:近7日≥10(不限)且近3日均≥5且加速≥3且涨≥3且剔尖峰+后3连续+非供给恢复｜"
                        "新品:前4日=0且有货未断+近3日均≥5+后3连续+末日不回吐"),
            "counts": {
                "size": len(size_rows), "speed": len(speed_rows), "new": len(new_rows),
                "size_tier": {t: sum(1 for x in size_rows if x["tier"] == t) for t, _ in
                              (("爆款", 0), ("旺销", 1), ("畅销", 2), ("起量", 3))},
                "speed_tier": {t: sum(1 for x in speed_rows if x["tier"] == t) for t in ("黑马", "猛涨", "起势")},
            },
            "on_shelf": on_shelf, "pool": pool,
            "snap_days": snap_days,
            "sold_out_excluded": sold_out_today,
            "pool_note": "当日售罄剔除后，近7日销量≥10件的品即热品池",
        },
        "size": size_rows, "speed": speed_rows, "new": new_rows,
    }

    out_json = os.path.join(a.outdir, "qpwy_爆卖_%s.json" % target.replace("-", ""))
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)

    # ---------------- 日志 + SQL 记录
    log_dir = os.path.join(a.outdir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    sql_path = a.sql_log or os.path.join(log_dir, "hot_sale_sql_%s.log" % stamp)
    with open(sql_path, "w", encoding="utf-8") as f:
        for kind, note, sql in db.sql_log:
            f.write("-- [%s] %s\n%s\n\n" % (kind, note, sql))
    log_path = os.path.join(log_dir, "hot_sale_%s.log" % stamp)
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(json.dumps({
            "ts": stamp, "day": target, "window": days,
            "on_shelf": on_shelf, "pool": pool, "sold_out_excluded": sold_out_today,
            "size": len(size_rows), "speed": len(speed_rows), "new": len(new_rows),
            "snap_days": snap_days, "elapsed_s": round(time.time() - t0, 1),
            "out_json": out_json, "sql_log": sql_path,
        }, ensure_ascii=False, indent=1))

    print(json.dumps({
        "day": target, "on_shelf": on_shelf, "pool": pool,
        "sold_out_excluded": sold_out_today,
        "size": len(size_rows), "speed": len(speed_rows), "new": len(new_rows),
        "size_tier": payload["meta"]["counts"]["size_tier"],
        "speed_tier": payload["meta"]["counts"]["speed_tier"],
        "elapsed_s": round(time.time() - t0, 1),
        "out_json": out_json, "log": log_path, "sql_log": sql_path,
    }, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
