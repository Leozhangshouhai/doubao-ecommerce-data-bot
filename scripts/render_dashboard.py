#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QPWY 每日看板出页器（唯一入口）
读四张表当日 JSON → 一页 dashboard.html（版式锁死，只许数字变）

用法：
  python3 scripts/render_dashboard.py --date 20261122 \
      --reports "/home/user/Doubao/chats/38449703070143746/AI电商数据洞察服务/04_样板案例/QPWY" \
      --out "../04_样板案例/QPWY/06_看板/dashboard.html"

约定：
- 只读 JSON，不连库、不重算口径；任何数字必须与当日 JSON 顶层统计相等
- 某张表 JSON 缺失 → 该卡片写「当日无数据（原因）」，照常出页，不造数
- 单文件、无外链（断网双击可开）
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

GRADE_ORDER = ["A", "B", "C", "D"]

CSS = """
:root{
  --title:#111827; --sub:#6B7280; --hdr:#1F2937; --line:#E5E7EB; --bg:#F8FAFC; --card:#fff;
  --A:#DC2626; --A-l:#FEE2E2; --A-t:#B91C1C;
  --B:#C2410C; --B-l:#FFEDD5; --B-t:#C2410C;
  --C:#A16207; --C-l:#FEF9C3; --C-t:#A16207;
  --D:#4B5563; --D-l:#F3F4F6; --D-t:#4B5563;
}
*{box-sizing:border-box}
html,body{margin:0;background:var(--bg);color:var(--title);
  font:15px/1.45 -apple-system,BlinkMacSystemFont,"PingFang SC","Hiragino Sans GB","Microsoft YaHei",sans-serif}
body{padding-bottom:40px}
.wrap{max-width:1080px;margin:0 auto;padding:16px 14px 24px}
h1{font-size:22px;margin:0;letter-spacing:.02em}
.meta{color:var(--sub);font-size:13px;margin-top:4px}
.rule{background:#FFF7ED;color:#9A3412;border:1px solid #FFEDD5;border-radius:10px;padding:9px 12px;font-size:13px;margin:12px 0 14px}
.kpis{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin:0 0 6px}
@media(min-width:720px){.kpis{grid-template-columns:repeat(5,1fr)}}
.kpi{background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px 13px 11px}
.kpi .n{font-size:26px;font-weight:700;line-height:1.1}
.kpi .l{font-size:12px;font-weight:600;margin-top:4px}
.kpi .d{font-size:11px;color:var(--sub);margin-top:2px}
.kpi.red{background:var(--A-l)} .kpi.red .n,.kpi.red .l{color:var(--A-t)}
.kpi.orange{background:var(--B-l)} .kpi.orange .n,.kpi.orange .l{color:var(--B-t)}
.grid{display:grid;gap:10px}
.grid>*{min-width:0}
@media(min-width:720px){.grid{grid-template-columns:1fr 1fr}}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px;overflow:hidden;min-width:0}
.card h2{margin:0 0 2px;font-size:16px}
.card .q{margin:0 0 8px;color:var(--sub);font-size:13px}
.concl{font-size:13px;color:var(--title);background:#F9FAFB;border-left:3px solid var(--hdr);padding:7px 10px;border-radius:0 8px 8px 0;margin:0 0 10px}
.numbers{display:flex;flex-wrap:wrap;gap:14px;margin:0 0 10px}
.numbers .box{min-width:72px}
.numbers .box b{display:block;font-size:24px;line-height:1.15}
.numbers .box span{font-size:12px;color:var(--sub)}
.dist{margin:2px 0 10px}
.dline{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:12px}
.dline .g{min-width:44px;text-align:center;border-radius:5px;font-weight:700;padding:1px 6px;white-space:nowrap}
.dline .track{display:block;flex:1;height:7px;background:#F3F4F6;border-radius:99px;overflow:hidden}
.dline .fill{display:block;height:100%}
.dline .num{width:46px;text-align:right;color:var(--sub);font-variant-numeric:tabular-nums}
.g.A{background:var(--A-l);color:var(--A-t)} .g.B{background:var(--B-l);color:var(--B-t)}
.g.C{background:var(--C-l);color:var(--C-t)} .g.D{background:var(--D-l);color:var(--D-t)}
.t3{margin:0;padding:0;list-style:none}
.t3 li{display:flex;gap:8px;align-items:baseline;padding:5px 0;border-top:1px solid var(--line);font-size:12.5px}
.t3 li:first-child{border-top:0}
.t3 .nm{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.t3 .sp{color:var(--sub);font-size:11.5px;max-width:52%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right}
.pri{display:inline-block;color:#fff;border-radius:6px;padding:1px 6px;font-size:11px;font-weight:600;background:var(--hdr)}
.pri.A{background:var(--A)} .pri.B{background:var(--B)} .pri.C{background:var(--C)} .pri.D{background:var(--D)}
.tag{display:inline-block;border:1px solid var(--line);border-radius:999px;padding:1px 7px;font-size:11px;color:var(--sub)}
.nodata{color:var(--sub);font-size:13px;padding:14px 0}
.foot{margin-top:20px;color:var(--sub);font-size:12px;line-height:1.7}
.foot b{color:var(--title)}
"""


def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def find_report(root, pattern, date):
    hits = []
    for d in ("03_报表-卖断货/日报", "02_报表-爆卖/日报", "05_报表-热品预判缺货/日报", "04_报表-供货缺货/日报"):
        p = os.path.join(root, d)
        if not os.path.isdir(p):
            continue
        for n in os.listdir(p):
            if re.search(pattern, n) and n.endswith(".json") and date in n:
                hits.append(os.path.join(p, n))
    return sorted(hits)[-1] if hits else None


def latest_date(root):
    cands = []
    for d, pat in (("03_报表-卖断货/日报", r"qpwy_卖断货_(\d{8})\.json"),
                   ("02_报表-爆卖/日报", r"qpwy_爆卖_(\d{8})\.json"),
                   ("05_报表-热品预判缺货/日报", r"qpwy_热品预判缺货_(\d{8})\.json"),
                   ("04_报表-供货缺货/日报", r"04-供货商快缺货_(\d{8})\.json")):
        p = os.path.join(root, d)
        if not os.path.isdir(p):
            continue
        for n in os.listdir(p):
            m = re.match(pat, n)
            if m:
                cands.append(m.group(1))
    return sorted(set(cands))[-1] if cands else None


def pct(x):
    try:
        return "%.1f%%" % (float(x) * 100)
    except Exception:
        return "—"


def num(x):
    try:
        f = float(x)
        return ("%d" % int(f)) if abs(f - int(f)) < 1e-9 else ("%.2f" % f)
    except (TypeError, ValueError):
        return "—"


def dist_block(items):
    """items = [(标签, 数值, 配色键 A/B/C/D), ...]"""
    total = sum(float(v or 0) for _, v, _ in items) or 1.0
    rows = []
    for label, v, ck in items:
        v = float(v or 0)
        w = max(2.0, (v / total) * 100) if v else 0
        rows.append(
            '<div class="dline"><span class="g %s">%s</span>'
            '<span class="track"><span class="fill" style="width:%.1f%%;background:var(--%s)"></span></span>'
            '<span class="num">%s</span></div>' % (ck, label, w, ck, num(v))
        )
    return '<div class="dist">%s</div>' % "".join(rows)


def li_grade(grade, text):
    return '<span class="pri %s">%s</span>' % (grade, text)


def li_tag(text):
    return '<span class="tag">%s</span>' % text


def esc(s):
    return (str(s if s is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def card(title, question, concl, numbers, dist_html, top_items, missing=None):
    if missing:
        body = '<div class="nodata">当日无数据（%s）</div>' % esc(missing)
    else:
        nums = "".join('<div class="box"><b>%s</b><span>%s</span></div>' % (esc(n), esc(l)) for n, l in numbers)
        body = (
            '<div class="concl">%s</div>' % concl +
            '<div class="numbers">%s</div>' % nums +
            (dist_html or "") +
            '<ul class="t3">%s</ul>' % "".join(top_items)
        )
    return ('<article class="card"><h2>%s</h2><p class="q">%s</p>%s</article>'
            % (esc(title), esc(question), body))


def build(date, root):
    paths = {
        "A": find_report(root, r"^qpwy_卖断货_", date),
        "B": find_report(root, r"^qpwy_爆卖_", date),
        "C": find_report(root, r"^qpwy_热品预判缺货_", date),
        "D": find_report(root, r"^04-供货商快缺货_", date),
    }
    data = {}
    for k, p in paths.items():
        data[k] = load(p) if p and os.path.exists(p) else None
    if not any(data.values()):
        return None, None

    # ---------- ① 卖断货 ----------
    d1 = data["A"]
    if d1:
        meta = d1.get("meta", {})
        total = meta.get("soldout_total") or len(d1.get("rows") or [])
        dist = meta.get("grade_dist") or {}
        rows = d1.get("rows") or []
        # 供商 TOP（A+B 条数）
        ab = {}
        for r in rows:
            if r.get("grade") in ("A", "B"):
                nm = r.get("provider_name") or "自营"
                ab[nm] = ab.get(nm, 0) + 1
        top_prov = sorted(ab.items(), key=lambda x: -x[1])[:3]
        watch = sorted([r for r in rows if r.get("grade") == "A"],
                       key=lambda r: -(r.get("max_month_amt") or 0))[:3]
        top_items = []
        for nm, n in top_prov:
            top_items.append('<li>%s<span class="nm">%s</span><span class="sp">%s 条</span></li>'
                             % (li_tag("A+B"), esc(nm), n))
        for r in watch:
            top_items.append('<li>%s<span class="nm">%s</span><span class="sp">连续缺 %s 天 · %s</span></li>'
                             % (li_grade("A", r.get("grade") or "A"), esc(r.get("name")),
                                esc(r.get("consec_days")), esc(r.get("provider_name") or "自营")))
        soldout_card = card("① 卖断货", "哪些品卖断货了",
                            "先看 A（当天售罄）+ B（连续缺货），C/D 是低销长尾，不用天天盯。",
                            [(total, "件售罄"),
                             ((dist.get("A") or 0) + (dist.get("B") or 0), "A+B 要处理"),
                             (dist.get("A") or 0, "A 当天售罄")],
                            dist_block([(g, dist.get(g) or 0, g) for g in GRADE_ORDER]), top_items)
        n1 = dict(total=total, dist=dist)
    else:
        soldout_card = card("① 卖断货", "哪些品卖断货了", "", [], "", [], "缺 qpwy_卖断货_%s.json" % date)
        n1 = dict(total=0, dist={})

    # ---------- ② 爆卖 ----------
    d2 = data["B"]
    if d2:
        c = (d2.get("meta") or {}).get("counts") or {}
        size, speed, new = c.get("size") or 0, c.get("speed") or 0, c.get("new") or 0
        stier = c.get("size_tier") or {}
        rows = sorted(d2.get("size") or [], key=lambda r: -(r.get("acc") or 0))[:3]
        tier_ck = {"爆款": "A", "旺销": "B", "畅销": "C", "起量": "D"}
        top_items = ['<li>%s<span class="nm">%s</span><span class="sp">近7日 %s 件 · 加速 %s× · %s</span></li>'
                     % (li_grade(tier_ck.get(r.get("tier"), "D"), r.get("tier") or "起量"), esc(r.get("name")),
                        num(r.get("qty_7d")), num(r.get("acc")), esc(r.get("provider_name") or "自营"))
                     for r in rows]
        burst_card = card("② 爆卖", "哪些品突然卖得远高于自己平时",
                          "规模榜看「够不够大」，速度榜看「涨得多急」；榜上品有断货风险 → 转 ③ 处理。",
                          [(size, "规模爆品"), (speed, "速度爆品"), (new, "新品起量")],
                          dist_block([("爆款", stier.get("爆款") or 0, "A"),
                                      ("旺销", stier.get("旺销") or 0, "B"),
                                      ("畅销", stier.get("畅销") or 0, "C"),
                                      ("起量", stier.get("起量") or 0, "D")]), top_items)
        n2 = dict(size=size, speed=speed, new=new, stier=stier)
    else:
        burst_card = card("② 爆卖", "哪些品突然卖得远高于自己平时", "", [], "", [], "缺 qpwy_爆卖_%s.json" % date)
        n2 = dict(size=0, speed=0, new=0, stier={})

    # ---------- ③ 热品预判缺货 ----------
    d3 = data["C"]
    if d3:
        cnt = d3.get("counts") or {}
        rows = sorted([r for r in (d3.get("rows") or []) if r.get("level") == "P0"],
                      key=lambda r: (r.get("cover") or 0))[:3] or \
               sorted(d3.get("rows") or [], key=lambda r: (r.get("cover") or 0))[:3]
        top_items = ['<li>%s<span class="nm">%s</span><span class="sp">可售 %s 天 · 日均 %s · 预估断 %s</span></li>'
                     % (li_grade("A", r.get("level") or "P0"), esc(r.get("name")),
                        num(r.get("cover")), num(r.get("daily")), esc(r.get("est_out")))
                     for r in rows]
        heat_card = card("③ 热品预判缺货", "哪些热品快断，得提前下单",
                         "P0 今天就可能断（≤1 天），P1 ≤2 天；这批是采购当天要处理的单。",
                         [(sum(cnt.values()) if cnt else len(d3.get("rows") or []), "预警总数"),
                          (cnt.get("P0") or 0, "P0 ≤1 天"), (cnt.get("P1") or 0, "P1 ≤2 天")],
                         dist_block([("P0", cnt.get("P0") or 0, "A"),
                                     ("P1", cnt.get("P1") or 0, "B"),
                                     ("P2", cnt.get("P2") or 0, "C"),
                                     ("P3", cnt.get("P3") or 0, "D")]), top_items)
        n3 = dict(counts=cnt)
    else:
        heat_card = card("③ 热品预判缺货", "哪些热品快断，得提前下单", "", [], "", [], "缺 qpwy_热品预判缺货_%s.json" % date)
        n3 = dict(counts={})

    # ---------- ④ 供货缺货 ----------
    d4 = data["D"]
    if d4 and (d4.get("suppliers") or []):
        meta = d4.get("meta") or {}
        sups = sorted(d4.get("suppliers") or [], key=lambda s: -(s.get("score") or 0))
        top_items = ['<li>%s<span class="nm">%s</span><span class="sp">关注分 %s · A缺 %s · 售罄 %s/%s</span></li>'
                     % (li_tag("No.%s" % s.get("rank")), esc(s.get("provider_name")), num(s.get("score")),
                        s.get("A"), s.get("sold"), s.get("on_shelf"))
                     for s in sups[:3]]
        supply_card = card("④ 供货缺货", "哪个供商快缺货",
                           "连续两个库存日都断才算；关注分 = 5×A+2×B+1×C+0.3×D，越高越先谈。",
                           [(len(sups), "上榜供商"),
                            (sups[0].get("A") or 0, "第一名 A 类缺口"),
                            (sups[0].get("score"), "头名关注分")],
                           "", top_items)
        n4 = dict(supplier_count=len(sups))
    else:
        supply_card = card("④ 供货缺货", "哪个供商快缺货", "", [], "", [], "缺 04-供货商快缺货_%s.json" % date)
        n4 = dict(supplier_count=0)

    # ---------- 顶部总览 ----------
    kpis = [
        ("red", n1.get("total") or 0, "件售罄", "① 卖断货"),
        ("orange", (n1.get("dist") or {}).get("A", 0) if n1.get("dist") else 0, "A 类当天售罄", "先处理这批"),
        ("", n2.get("size") or 0, "规模爆品", "② 爆卖·近7日"),
        ("", sum((n3.get("counts") or {}).values()), "缺货预警", "③ 含 P0 %s" % ((n3.get("counts") or {}).get("P0") or 0)),
        ("", n4.get("supplier_count") or 0, "快断供商", "④ 连续两天"),
    ]
    kpi_html = "".join(
        '<div class="kpi %s"><div class="n">%s</div><div class="l">%s</div><div class="d">%s</div></div>'
        % (cls, num(v), esc(l), esc(d)) for cls, v, l, d in kpis
    )

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    day_cn = "%s-%s-%s" % (date[0:4], date[4:6], date[6:8])

    html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="color-scheme" content="light">
<title>QPWY 今日看板 · {day}</title>
<style>{css}</style>
</head>
<body>
<div class="wrap">
  <h1>QPWY 今日看板</h1>
  <div class="meta">数据日 {day}（库存快照） · 生成 {now} · 四张固定报表核心数汇总，明细看 XLSX</div>
  <div class="rule">口径提醒：库存快照每天 23:50 写当天，早上跑出来的数据日 = 昨天；A 类=当天售罄，B 类=连续缺货；供商列「自营」= 没有供商的品（内部补货）。</div>
  <div class="kpis">{kpis}</div>
  <div class="grid">
    {c1}
    {c2}
    {c3}
    {c4}
  </div>
  <div class="foot">
    <b>数据来源</b>：① 卖断货 / ② 爆卖 / ③ 热品预判缺货 / ④ 供货缺货 当日 JSON（只读库取数，JSON 留档不交付）。<br>
    <b>怎么用</b>：一屏看风险，要名单/要发给下属就开同日的 XLSX 底稿。<br>
    <b>生成器</b>：scripts/render_dashboard.py（版式锁死，禁止改配色与结构） · 数据日 {day} · {now}
  </div>
</div>
</body>
</html>
""".format(css=CSS, day=day_cn, now=now, kpis=kpi_html,
           c1=soldout_card, c2=burst_card, c3=heat_card, c4=supply_card)
    return html, dict(n1=n1, n2=n2, n3=n3, n4=n4)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=None, help="YYYYMMDD，缺省=四表里最新共有日期")
    ap.add_argument("--reports", default="/home/user/Doubao/chats/38449703070143746/AI电商数据洞察服务/04_样板案例/QPWY")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    date = args.date or latest_date(args.reports)
    if not date:
        print("[FATAL] 找不到任何日报 JSON，无法出看板")
        sys.exit(2)
    if not re.match(r"^\d{8}$", date):
        print("[FATAL] --date 必须是 YYYYMMDD")
        sys.exit(2)

    html, stats = build(date, args.reports)
    if html is None:
        print("[FATAL] %s 没有任何一张表的数据，看板不出（避免交一张空页）" % date)
        sys.exit(2)

    out = args.out or os.path.join(args.reports, "06_看板", "dashboard.html")
    outdir = os.path.dirname(os.path.abspath(out))
    os.makedirs(outdir, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    archive = os.path.join(outdir, "dashboard_%s.html" % date)
    with open(archive, "w", encoding="utf-8") as f:
        f.write(html)

    # ---------- 回读自检 ----------
    size = os.path.getsize(out)
    txt = open(out, encoding="utf-8").read()
    errs = []
    if "None" in txt or "NaN" in txt or "nan" in txt:
        errs.append("页面出现 None/NaN")
    for must in ("① 卖断货", "② 爆卖", "③ 热品预判缺货", "④ 供货缺货"):
        if must not in txt:
            errs.append("缺卡片：%s" % must)
    if re.search(r'(src|href)="(https?:)?//', txt):
        errs.append("存在外链资源，断网打不开")
    n1 = stats["n1"]
    if n1.get("total") and str(n1["total"]) not in txt:
        errs.append("售罄总数 %s 未出现在页面" % n1["total"])
    n3 = stats["n3"].get("counts") or {}
    if n3 and str(sum(n3.values())) not in txt:
        errs.append("预警总数与页面不一致")
    warns = []
    partial = txt.count("当日无数据（")
    if partial:
        warns.append("有 %d 张表当日无数据，已按「空卡片」降级出页" % partial)
    if size < (6000 if partial else 10000):
        errs.append("文件过小（%d 字节），疑似渲染失败" % size)

    print("[OK] 数据日 %s" % date)
    print("     售罄 %s｜A %s / B %s / C %s / D %s" % (
        n1.get("total"),
        (n1.get("dist") or {}).get("A"), (n1.get("dist") or {}).get("B"),
        (n1.get("dist") or {}).get("C"), (n1.get("dist") or {}).get("D")))
    print("     爆卖 规模 %s / 速度 %s / 新品 %s" % (stats["n2"].get("size"), stats["n2"].get("speed"), stats["n2"].get("new")))
    print("     预警 合计 %s（P0 %s / P1 %s / P2 %s / P3 %s）" % (sum(n3.values()) if n3 else 0,
                                                          n3.get("P0") or 0, n3.get("P1") or 0,
                                                          n3.get("P2") or 0, n3.get("P3") or 0))
    print("     供商 上榜 %s" % stats["n4"].get("supplier_count"))
    print("     输出 %s（%d 字节）＋归档 %s" % (out, size, os.path.basename(archive)))
    if errs:
        print("[FAIL] 自检未通过：")
        for e in errs:
            print("       - %s" % e)
        sys.exit(1)
    for w in warns:
        print("[WARN] %s" % w)
    print("[SELF-CHECK] 通过：四卡齐全 / 数字与 JSON 一致 / 无 None NaN / 无外链")


if __name__ == "__main__":
    main()
