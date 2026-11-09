#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
③ 热品预判缺货 取数器（只读）。

定位：给采购的**提前预警**——今天就可能断的品，今天下单，T+1 到货还接得上；
      不是"已经断了"（已断的归 ① 卖断货 / ④ 供货缺货）。

！！！ 关键设计原则 ！！！
③ **不是** ② 爆卖的子集。稳定日销 300 件、库存只够 2 天的品不"暴涨"，进不了 ②，
但它马上要断，必须出现在 ③。所以这里自己从"热品池"取数，不从 ② 的名单里筛。
② 只用于打一个标记「是否爆卖榜」（= 又涨又快断，最高优先级）。

热品池 = 近 7 日销量 ≥ min_q7 件 且 当日未断货 的在售品
③ 名单 = 热品池 ∩ 可售天数 ≤ max_cover_days
可售天数 = 当前库存 ÷ 日均；日均 = max(近3日均, 近7日日均)（保守，宁可早报）

产出 JSON，供 scripts/render_predict.py 出表。
口径见 sql_templates/predict.yaml。

用法：
  python3 report_predict.py --config /path/ecommerce_bot.yaml \
      --outdir "../04_样板案例/QPWY/05_报表-热品预判缺货/日报" [--date YYYY-MM-DD]
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from report_hot_sale import DB   # 复用同一套跳板机隧道与只读查询封装


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True, help="ecommerce_bot.yaml 路径")
    ap.add_argument("--outdir", default="./report_out")
    ap.add_argument("--date", default=None, help="目标库存日 YYYY-MM-DD，默认取最新快照日")
    ap.add_argument("--hot-json", default=None, help="② 爆卖的 JSON，用于标注「是否爆卖榜」；缺省自动在同目录找")
    a = ap.parse_args()

    import yaml
    with open(a.config, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    tables = cfg["tables"]
    so = cfg["caliber"]["soldout"]
    pk = cfg["caliber"].get("predict", {})
    min_q7 = float(pk.get("min_q7", 10))
    max_cov = float(pk.get("max_cover_days", 15))
    lv = pk.get("levels", {}) or {}
    LEVEL_TH = [("P0", float(lv.get("P0", 1))), ("P1", float(lv.get("P1", 2))),
                ("P2", float(lv.get("P2", 5))), ("P3", float(lv.get("P3", 15)))]
    flow_cats = list(so["flow_cats"])

    os.makedirs(a.outdir, exist_ok=True)
    db = DB(cfg)
    t0 = time.time()
    db.connect()
    try:
        # 1) 目标日
        tgt = db.q("order", "SELECT MAX(snapshot_date) FROM %s WHERE snapshot_date IS NOT NULL"
                   % tables["inventory_sale"], "目标库存日")
        target = str(tgt[0][0]) if tgt and tgt[0][0] else None
        if a.date:
            target = a.date
        if not target:
            raise SystemExit("[FATAL] inventory_daily_sale 无任何快照，终止")
        print("[1/6] 目标日 %s" % target)

        td = datetime.strptime(target, "%Y-%m-%d").date()
        days = [str(td - timedelta(days=i)) for i in range(6, -1, -1)]
        print("[2/6] 近7日窗口 %s ~ %s" % (days[0], days[-1]))

        # 2) 上架商品（与 ①④ 同口径：status=2 & sale_type=1）
        on = db.q("goods", "SELECT item_id, name, provider_id FROM %s WHERE status=%s AND sale_type=%s"
                  % (tables["item_sale"], int(so["on_sale_status"]), int(so["sale_type"])), "上架商品")
        items = {}
        for iid, nm, pid in on:
            it = items.setdefault(int(iid), {"name": nm, "pids": []})
            it["pids"].append(int(pid or 0))
        for it in items.values():
            it["pids"] = sorted(set(it["pids"]))
            it["pid"] = 0 if 0 in it["pids"] else it["pids"][0]
        print("[3/6] 上架商品 %d" % len(items))

        # 3) 供商名称（去「（云仓）」后缀，与 ① 一致）
        prov = db.q("goods", "SELECT provider_id, provider_name FROM %s WHERE status=1"
                    % tables["item_provider"], "供商名称")
        import re as _re
        prov_name = {}
        for r in prov:
            nm = _re.sub(r"[（(]\s*云仓\s*[)）]", "", str(r[1] or "")).strip()
            prov_name[str(r[0])] = nm

        # 4) 类目
        g = db.q("goods", "SELECT item_id, top_cat FROM %s" % tables["goods"], "商品类目")
        topcat = {int(r[0]): r[1] for r in g}
        ids = set(v for v in topcat.values() if v)
        cat_name = {}
        if ids:
            ph = ",".join(["'%s'" % str(x) for x in ids])
            cats = db.q("goods", "SELECT id, name FROM %s WHERE id IN (%s)"
                        % (tables["item_category"], ph), "类目名称")
            cat_name = {r[0]: r[1] for r in cats}

        # 5) 近7日逐日销量
        sql = pk.get("sales_sql") or """
SELECT od.sku_code, DATE(FROM_UNIXTIME(oh.order_time)) AS d, SUM(od.quantity) AS q
FROM {order_detail} od JOIN {order_head} oh ON od.order_code = oh.order_code
WHERE oh.valid = 1 AND oh.order_time >= {ts0}
GROUP BY od.sku_code, DATE(FROM_UNIXTIME(oh.order_time))"""
        ts0 = int(datetime.strptime(days[0], "%Y-%m-%d").replace(tzinfo=None).timestamp()) - 8 * 3600
        sale = db.q("order", sql.format(order_detail=tables["order_detail"],
                                        order_head=tables["order_head"], ts0=ts0), "逐日销量")
        daily = {}
        for code, d, q in sale:
            try:
                iid = int(code)
            except Exception:
                continue
            k = str(d)[:10]
            if k in days:
                daily.setdefault(iid, {})[k] = float(q or 0)
        print("[4/6] 窗口内有销量商品 %d" % len(daily))

        # 6) 目标日库存（SUM 多 zone，按 item+供商聚合）
        inv_sql = pk.get("stock_sql") or """
SELECT item_id, provider_id, SUM(total_qty) AS tq, MAX(is_sold_out) AS so
FROM {inventory_sale} WHERE snapshot_date = '{day}'
GROUP BY item_id, provider_id"""
        inv = db.q("order", inv_sql.format(inventory_sale=tables["inventory_sale"], day=target), "目标日库存")
        stock_map = {}
        for iid, pid, tq, sflag in inv:
            stock_map[(int(iid), int(pid or 0))] = (float(tq or 0), int(sflag or 0))
        print("[5/6] 目标日库存键 %d" % len(stock_map))

        # ② 爆卖榜名单（可选）
        hot_ids = set()
        hj = a.hot_json or os.path.join(a.outdir, "qpwy_爆卖_%s.json" % target.replace("-", ""))
        if os.path.isfile(hj):
            try:
                hd = json.load(open(hj, encoding="utf-8"))
                for k in ("size", "speed", "rows"):
                    v = hd.get(k)
                    if isinstance(v, list):
                        for r in v:
                            if isinstance(r, dict) and r.get("item_id") is not None:
                                hot_ids.add(int(r["item_id"]))
                print("[6/6] 叠加 ② 爆卖榜 %d 个：%s" % (len(hot_ids), os.path.basename(hj)))
            except Exception as e:
                print("[6/6] ② 榜单读取失败（不影响 ③ 主表）：%s" % e)
        else:
            print("[6/6] 未找到 ② 的 JSON，「是否爆卖榜」列留空")

        # ---------------- 计算 ----------------
        rows, pool, n_soldout = [], 0, 0
        for iid, meta in items.items():
            series = [daily.get(iid, {}).get(d, 0.0) for d in days]
            q7 = sum(series)
            if q7 < min_q7:
                continue
            pid = meta["pid"]
            st = stock_map.get((iid, pid))
            if st is None:
                continue
            stock, sold = st
            # 上架供商（item_sale status=2 那条）当日那一行的 is_sold_out 就是 ① 的判据。
            # 两表必须同一个断货判据，否则会出现「① 说已售罄、③ 说还能卖几天」的矛盾（实测 12 个）。
            if sold:
                n_soldout += 1
                continue
            pool += 1
            d7 = q7 / 7.0
            d3 = sum(series[-3:]) / 3.0
            aggr = max(d3, d7)               # 保守：宁可早报
            if aggr <= 0:
                continue
            cov = stock / aggr
            if cov > max_cov:
                continue
            level = "P3"
            for name, th in LEVEL_TH:
                if cov <= th:
                    level = name
                    break
            import math as _m
            est_days = max(int(_m.ceil(cov)), 1)   # 0.09 天也算「今天就会断」，不写成 0
            c = topcat.get(iid)
            fname = cat_name.get(str(c)) if c else "?"
            notes = []
            mx = max(series)
            rest = [v for v in series if v != mx]
            if mx >= 10 and rest and sum(rest) / len(rest) > 0 and mx > sum(rest) / len(rest) * 5:
                notes.append("含单日尖峰 %d 件" % mx)
            if sum(series[-3:]) <= 0:
                notes.append("近3日无销量，日均取自更早几日，需核对是否已停售")
            rows.append({
                "item_id": iid,
                "name": meta["name"],
                "provider_id": pid,
                "provider_name": prov_name.get(str(pid), "") if pid else "",
                "fname": fname,
                "flow": "流通" if fname in flow_cats else "非流通",
                "q7": q7, "d3": round(d3, 2), "d7": round(d7, 2), "daily": round(aggr, 2),
                "stock": round(stock, 2),
                "cover": round(cov, 2),
                "level": level,
                "est_out": str(td + timedelta(days=est_days)),
                "order_by": str(td + timedelta(days=max(est_days - 1, 1))),
                "in_hot": iid in hot_ids,
                "note": "；".join(notes),
                "series": [round(v, 1) for v in series],
            })

        # 排序：按档位分段（P0→P3）；段内正常品在前、带异常标注的排末尾（不剔除）；再按可售天数升序、近7日降序
        lv_rank = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}
        rows.sort(key=lambda r: (lv_rank[r["level"]], 1 if r["note"] else 0, r["cover"], -r["q7"]))
        cnt = {}
        for r in rows:
            cnt[r["level"]] = cnt.get(r["level"], 0) + 1

        out = {"day": target, "win": days, "pool": pool, "counts": cnt,
               "thresholds": {"min_q7": min_q7, "max_cover_days": max_cov,
                              "levels": {k: v for k, v in LEVEL_TH}},
               "rows": rows}
        jp = os.path.join(a.outdir, "qpwy_热品预判缺货_%s.json" % target.replace("-", ""))
        json.dump(out, open(jp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

        print("\n[完成] 热品池 %d 个 → 预警 %d 个 ｜ P0=%d P1=%d P2=%d P3=%d ｜ 当日已断已排除 %d"
              % (pool, len(rows), cnt.get("P0", 0), cnt.get("P1", 0), cnt.get("P2", 0), cnt.get("P3", 0), n_soldout))
        print("JSON: %s" % jp)
        print("耗时 %.1fs" % (time.time() - t0))
    finally:
        db.close()


if __name__ == "__main__":
    main()
