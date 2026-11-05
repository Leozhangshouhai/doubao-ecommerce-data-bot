#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
供货商快缺货（报表 ④）。只读。

口径与卖断货生成器相同的部分：上架 status=2 sale_type=1、售罄=该供商 is_sold_out=1、
月销=valid=1 的 price*quantity、大分类=flow_cats、等级函数与 report_soldout.grade_of 相同。
本表额外规则：最近两个库存日都售罄才计入；关注分=5A+2B+1C+0.3D；取前 20 家。

用法：
  python3 report_supply_short.py --self-check
  python3 report_supply_short.py --config /path/ecommerce_bot.yaml [--date YYYY-MM-DD] [--outdir ./report_out]

库存日不足 2 个时退出，不出表。不读 Redis，不读 INVENTORY_QUANTITY。
"""
import argparse
import json
import os
import re
import socket
import sys
import threading
import time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

YAHEI = "Microsoft YaHei"
C = {
    "title": "111827", "subtitle": "6B7280", "hdr": "1F2937", "hot": "DC2626",
    "A_l": "FEE2E2", "A_t": "B91C1C",
    "B_l": "FFEDD5", "B_t": "C2410C",
    "C_l": "FEF9C3", "C_t": "A16207",
    "D_l": "F3F4F6", "D_t": "4B5563",
}
WEIGHT_DEFAULT = {"A": 5.0, "B": 2.0, "C": 1.0, "D": 0.3}
TOP_N_DEFAULT = 20
FLOW_CANON = ["中外名酒", "方便速食", "牛奶乳品", "米面粮油", "饮料及水"]
GRADE_HOW = {
    "A": ("5", "高销", "流通品最高一个月销售额超过 1 万；其余商品超过 2 千"),
    "B": ("2", "中销", "流通品 1 千到 1 万；其余商品 1 千到 2 千"),
    "C": ("1", "一般", "最高一个月销售额 500 到 1 千"),
    "D": ("0.3", "低销", "最高一个月销售额不到 500"),
}
RULE_TEXT = "只统计连续 2 个库存日都售罄的商品。关注分 = 5×A + 2×B + 1×C + 0.3×D。今天才断的不计。"
SIDE = Side(style="thin", color="E5E7EB")
GRID = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)
NOB = Border()


def grade_of(max_amt, flow, th):
    """与 scripts/report_soldout.py 的 grade_of 相同。1000 整是 B，500 整是 C。"""
    if max_amt > th[flow]["A"]:
        return "A"
    if max_amt >= th[flow]["B"]:
        return "B"
    if max_amt >= th[flow]["C"]:
        return "C"
    return "D"


def clean_name(nm):
    nm = str(nm or "")
    for token in ("（云仓）", "(云仓）", "（云仓)", "(云仓)"):
        nm = nm.replace(token, "")
    return nm.strip()


def score_of(counts, weights):
    raw = sum(float(weights[g]) * int(counts[g]) for g in "ABCD")
    return round(raw + 1e-9, 1)


def points_of(grade, weights):
    return float(weights[grade])


def _both_providers(older_out, newer_out):
    """每个商品在两个库存日都售罄的快照供商。供商 0 不算。"""
    older, newer = {}, {}
    for item, pid in older_out:
        if int(pid) != 0:
            older.setdefault(int(item), set()).add(int(pid))
    for item, pid in newer_out:
        if int(pid) != 0:
            newer.setdefault(int(item), set()).add(int(pid))
    both = {}
    for item, pids in older.items():
        hit = pids & newer.get(item, set())
        if hit:
            both[item] = hit
    return both


def select_skus(on_shelf, older_out, newer_out, prov_name=None, zero_rows=None):
    """两天都售罄才返回。

    上架供商不是 0：只认这个供商的快照行。
    上架供商是 0：供商 0 行还有货（is_sold_out=0）→ 商品实际有货，不算售罄、不归属；
                  否则快照里恰好一个供商两天都售罄时，归到该供商；
                  多个快照供商都断着，不猜，这件不归任何一家。
    """
    prov_name = prov_name or {}
    zero_rows = zero_rows or {}
    both = _both_providers(older_out, newer_out)
    sold = []
    shelf = []
    fallback = 0
    for it in on_shelf:
        item = int(it["item_id"])
        pid = int(it["provider_id"] or 0)
        row = dict(it)
        if pid != 0:
            shelf.append(row)
            if (item, pid) in older_out and (item, pid) in newer_out:
                sold.append(dict(row))
            continue
        zrows = zero_rows.get(item)
        if zrows and any(int(v) == 0 for v in zrows.values()):
            # 供商 0 行还有货：商品没断，与卖断货同判据（卖断货读的是 (商品, 上架供商) 行）
            shelf.append(row)
            continue
        owners = both.get(item, set())
        if len(owners) != 1:
            shelf.append(row)
            continue
        owner = next(iter(owners))
        row["provider_id"] = owner
        row["provider_name"] = prov_name.get(owner, str(owner))
        shelf.append(row)
        sold.append(dict(row))
        fallback += 1
    return sold, shelf, fallback


def rank_suppliers(on_shelf, sold_skus, weights, top_n):
    shelf_n = {}
    for it in on_shelf:
        pid = int(it["provider_id"])
        shelf_n[pid] = shelf_n.get(pid, 0) + 1
    bucket = {}
    for sku in sold_skus:
        pid = int(sku["provider_id"])
        b = bucket.setdefault(pid, {
            "provider_id": pid,
            "provider_name": sku["provider_name"],
            "A": 0, "B": 0, "C": 0, "D": 0,
            "skus": [],
        })
        b[sku["grade"]] += 1
        b["skus"].append(sku)
    rows = []
    for pid, b in bucket.items():
        sold = b["A"] + b["B"] + b["C"] + b["D"]
        if sold < 1:
            continue
        shelf = shelf_n.get(pid, sold)
        rows.append({
            "provider_id": pid,
            "provider_name": b["provider_name"],
            "score": score_of(b, weights),
            "A": b["A"], "B": b["B"], "C": b["C"], "D": b["D"],
            "on_shelf": shelf,
            "sold": sold,
            "rate": sold / float(shelf) if shelf else 0.0,
            "skus": b["skus"],
        })
    rows.sort(key=lambda r: (-r["score"], -r["sold"], r["provider_name"]))
    rows = rows[:top_n]
    for i, r in enumerate(rows, 1):
        r["rank"] = i
        r["skus"].sort(key=lambda s: ("ABCD".index(s["grade"]), int(s["item_id"])))
    return rows


def yh(size, bold=False, color="111827"):
    return Font(name=YAHEI, size=size, bold=bold, color=color)


def solid(h):
    return PatternFill("solid", fgColor=h)


def none_fill():
    return PatternFill(fill_type=None)


def paint(cell, size, bold, color, fill, border, horizontal):
    cell.font = yh(size, bold, color)
    cell.fill = fill
    cell.border = border
    cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True)


def write_legend(ws, last_col, title, subtitle, flow_names):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    paint(ws.cell(1, 1, title), 16, True, C["title"], none_fill(), NOB, "left")
    ws.row_dimensions[1].height = 32
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    paint(ws.cell(2, 1, subtitle), 10, False, C["subtitle"], none_fill(), NOB, "left")
    ws.row_dimensions[2].height = 22
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
    paint(ws.cell(3, 1, "积分"), 10, True, C["title"], none_fill(), NOB, "center")
    ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=last_col)
    paint(ws.cell(3, 3, RULE_TEXT), 10, False, C["title"], none_fill(), NOB, "left")
    ws.row_dimensions[3].height = 22
    heads = ["档", "每件几分", "含义", "怎么划分（按今年卖得最高的那一个月）"]
    for col, text in enumerate(heads, 1):
        paint(ws.cell(4, col, text), 10, True, C["subtitle"], none_fill(), NOB, "center" if col < 4 else "left")
    if last_col > 4:
        ws.merge_cells(start_row=4, start_column=4, end_row=4, end_column=last_col)
    ws.row_dimensions[4].height = 18
    for i, g in enumerate("ABCD"):
        r = 5 + i
        pts, meaning, how = GRADE_HOW[g]
        bg, fg = C[g + "_l"], C[g + "_t"]
        vals = [g, pts, meaning, how]
        for col, text in enumerate(vals, 1):
            color = fg if col == 1 else C["title"]
            paint(ws.cell(r, col, text), 11, col == 1, color, solid(bg), GRID, "center" if col < 4 else "left")
        if last_col > 4:
            ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=last_col)
        ws.row_dimensions[r].height = 20
    note = "流通品指：" + "、".join(flow_names) + "。其余商品按「其余」那一档划分。"
    ws.merge_cells(start_row=9, start_column=1, end_row=9, end_column=last_col)
    paint(ws.cell(9, 1, note), 9, False, C["subtitle"], none_fill(), NOB, "left")
    ws.row_dimensions[9].height = 18


def page_setup(ws):
    ws.freeze_panes = "A11"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 120
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.print_title_rows = "1:10"
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = "关注分 = 5×A + 2×B + 1×C + 0.3×D · 只计连续 2 天售罄"
    ws.oddFooter.center.font = "Microsoft YaHei"
    ws.oddFooter.center.size = 9


def build_xlsx(suppliers, older, newer, flow_names, weights, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "前20供商"
    n = len(suppliers)
    write_legend(
        ws, 10, "供货商快缺货 · 前 20",
        "库存日 %s 与 %s 连续两天都售罄｜关注分最高的 %d 家｜这一页自己能看懂" % (older, newer, n),
        flow_names,
    )
    headers = ["排名", "供商", "关注分", "A 缺", "B 缺", "C 缺", "D 缺", "上架数", "售罄数", "售罄占比"]
    for col, text in enumerate(headers, 1):
        paint(ws.cell(10, col, text), 11, True, "FFFFFF", solid(C["hdr"]), GRID, "center")
    ws.row_dimensions[10].height = 22
    for i, row in enumerate(suppliers):
        r = 11 + i
        vals = [row["rank"], row["provider_name"], row["score"], row["A"], row["B"], row["C"], row["D"],
                row["on_shelf"], row["sold"], row["rate"]]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(r, col, val)
            color, fill, bold, size = C["title"], none_fill(), False, 11
            if col == 2:
                bold, size = True, 12
            elif col == 3:
                color, bold, size = C["hot"], True, 14
                cell.number_format = "0.0"
            elif col == 4:
                if row["A"] > 0:
                    color, fill, bold, size = C["A_t"], solid(C["A_l"]), True, 13
                else:
                    color, fill, bold, size = C["D_t"], solid(C["D_l"]), True, 13
            elif col == 10:
                cell.number_format = "0.0%"
            horizontal = "left" if col == 2 else "center"
            paint(cell, size, bold, color, fill, GRID, horizontal)
        ws.row_dimensions[r].height = 26
    last = 10 + max(n, 1)
    ws.auto_filter.ref = "A10:J%d" % (10 + n if n else 10)
    for col, width in enumerate([8, 36, 12, 10, 10, 10, 10, 12, 12, 12], 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    page_setup(ws)
    ws.auto_filter.ref = "A10:J%d" % last if n == 0 else "A10:J%d" % (10 + n)

    ws2 = wb.create_sheet("前20缺了什么")
    write_legend(
        ws2, 5, "前 20 家连续缺了什么",
        "只列出这 %d 家名下、连续两天都售罄的商品｜按供商排名，同供商 A→D｜这一页自己能看懂" % n,
        flow_names,
    )
    for col, text in enumerate(["供商", "档", "计入关注分", "商城商品码", "商品名"], 1):
        paint(ws2.cell(10, col, text), 11, True, "FFFFFF", solid(C["hdr"]), GRID, "center")
    ws2.row_dimensions[10].height = 22
    r = 11
    for sup in suppliers:
        ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        paint(ws2.cell(r, 1, sup["provider_name"]), 12, True, "FFFFFF", solid(C["hdr"]), GRID, "left")
        ws2.row_dimensions[r].height = 24
        r += 1
        for sku in sup["skus"]:
            g = sku["grade"]
            pts = points_of(g, weights)
            vals = [sup["provider_name"], g, pts, str(sku["item_id"]), sku["name"]]
            for col, val in enumerate(vals, 1):
                cell = ws2.cell(r, col, val)
                color, fill, bold = C["title"], none_fill(), g == "A"
                horizontal = "left" if col in (1, 5) else "center"
                if col == 2:
                    color, fill, bold = C[g + "_t"], solid(C[g + "_l"]), True
                elif col == 3:
                    color = C[g + "_t"]
                    bold = g == "A"
                    cell.number_format = "0.0" if g == "D" else "0"
                paint(cell, 11, bold, color, fill, GRID, horizontal)
            ws2.row_dimensions[r].height = 22
            r += 1
    for col, width in enumerate([36, 8, 14, 16, 52], 1):
        ws2.column_dimensions[get_column_letter(col)].width = width
    page_setup(ws2)
    wb.save(path)
    return path


def verify_workbook(path, suppliers):
    wb = load_workbook(path)
    errors = []
    if wb.sheetnames != ["前20供商", "前20缺了什么"]:
        errors.append("sheet 名不对：%s" % wb.sheetnames)
    for name in wb.sheetnames:
        ws = wb[name]
        if ws["A1"].font.name != YAHEI or ws["A1"].font.color.rgb != "00111827":
            errors.append("%s 标题字体或颜色不对" % name)
        if ws["A1"].fill.fgColor is not None and ws["A1"].fill.fgColor.rgb not in ("00000000", None):
            errors.append("%s 标题不该有色条" % name)
        if RULE_TEXT not in str(ws["C3"].value):
            errors.append("%s 缺少积分规则" % name)
        if "流通品指：" not in str(ws["A9"].value):
            errors.append("%s 缺少 ABCD 划分说明" % name)
        if ws["A5"].fill.fgColor.rgb != "00FEE2E2" or ws["A5"].font.color.rgb != "00B91C1C":
            errors.append("%s A 档颜色不是卖断货红" % name)
        if ws["A6"].fill.fgColor.rgb != "00FFEDD5":
            errors.append("%s B 档颜色不是卖断货橙" % name)
        if ws["A7"].fill.fgColor.rgb != "00FEF9C3":
            errors.append("%s C 档颜色不是卖断货黄" % name)
        if ws["A8"].fill.fgColor.rgb != "00F3F4F6":
            errors.append("%s D 档颜色不是卖断货灰" % name)
        if ws["A10"].fill.fgColor.rgb != "001F2937":
            errors.append("%s 表头不是卖断货深灰" % name)
        blob = " ".join(str(c.value) for row in ws.iter_rows(max_row=ws.max_row, max_col=ws.max_column) for c in row)
        if "云仓" in blob:
            errors.append("%s 供商名仍带云仓" % name)
    ws = wb["前20供商"]
    if suppliers and ws["C11"].font.color.rgb != "00DC2626":
        errors.append("关注分不是今日必盯红")
    names1 = []
    for i in range(len(suppliers)):
        names1.append(ws.cell(11 + i, 2).value)
    if names1 != [s["provider_name"] for s in suppliers]:
        errors.append("第一页供商顺序与结果不一致")
    ws2 = wb["前20缺了什么"]
    bands = []
    r = 11
    while r <= ws2.max_row:
        val = ws2.cell(r, 1).value
        grade = ws2.cell(r, 2).value
        if val and grade not in list("ABCD"):
            bands.append(val)
        r += 1
    if bands != [s["provider_name"] for s in suppliers]:
        errors.append("第二页供商与第一页不一致")
    if errors:
        raise SystemExit("[FATAL] 样式回读未过：\n- " + "\n- ".join(errors))


def flow_label(flow_cats):
    if set(flow_cats) == set(FLOW_CANON):
        return list(FLOW_CANON)
    return list(flow_cats)


def weights_from(cfg):
    block = ((cfg.get("caliber") or {}).get("supply_short") or {})
    w = dict(WEIGHT_DEFAULT)
    given = block.get("weight") or {}
    for g in "ABCD":
        if g in given:
            w[g] = float(given[g])
    top_n = int(block.get("top_n") or TOP_N_DEFAULT)
    return w, top_n


def rule_text(weights):
    global RULE_TEXT
    def num(g):
        v = float(weights[g])
        return str(int(v)) if v == int(v) else ("%s" % v).rstrip("0").rstrip(".")
    RULE_TEXT = "只统计连续 2 个库存日都售罄的商品。关注分 = %s×A + %s×B + %s×C + %s×D。今天才断的不计。" % (
        num("A"), num("B"), num("C"), num("D"))
    GRADE_HOW["A"] = (num("A"), "高销", GRADE_HOW["A"][2])
    GRADE_HOW["B"] = (num("B"), "中销", GRADE_HOW["B"][2])
    GRADE_HOW["C"] = (num("C"), "一般", GRADE_HOW["C"][2])
    GRADE_HOW["D"] = (num("D"), "低销", GRADE_HOW["D"][2])


def self_check():
    th = {"流通": {"A": 10000, "B": 1000, "C": 500}, "非流通": {"A": 2000, "B": 1000, "C": 500}}
    cases = [
        (10000, "流通", "B"), (10000.01, "流通", "A"),
        (2000, "非流通", "B"), (2000.01, "非流通", "A"),
        (1000, "流通", "B"), (500, "非流通", "C"), (499.99, "流通", "D"),
    ]
    for amt, flow, expect in cases:
        got = grade_of(amt, flow, th)
        if got != expect:
            raise SystemExit("[FATAL] 等级边界 %s %s 得到 %s，应为 %s" % (flow, amt, got, expect))
    on_shelf = []
    sold = []

    def add(pid, name, item_id, grade, both=True):
        on_shelf.append({"item_id": item_id, "name": "示例%d" % item_id, "provider_id": pid, "provider_name": name})
        if both:
            sold.append({"item_id": item_id, "name": "示例%d" % item_id, "provider_id": pid,
                         "provider_name": name, "grade": grade})
    for i, g in enumerate(["A"] * 4 + ["B"] * 3 + ["C"] * 2 + ["D"] * 10):
        add(1, "北京飞箭商贸有限公司", 1000 + i, g)
    add(1, "北京飞箭商贸有限公司", 1099, "A", both=False)
    on_shelf.extend([{"item_id": 2000 + i, "name": "在架%d" % i, "provider_id": 1, "provider_name": "北京飞箭商贸有限公司"} for i in range(86 - 20)])
    for i, g in enumerate(["B"] + ["C"] * 2 + ["D"] * 36):
        add(2, "北京德恒伟业商贸有限公司", 3000 + i, g)
    add(3, "只有一天的供商", 4000, "A", both=False)
    rows = rank_suppliers(on_shelf, sold, WEIGHT_DEFAULT, 20)
    by_name = {r["provider_name"]: r for r in rows}
    if abs(by_name["北京飞箭商贸有限公司"]["score"] - 31.0) > 0.05:
        raise SystemExit("[FATAL] 飞箭关注分应为 31，得到 %s" % by_name["北京飞箭商贸有限公司"]["score"])
    if by_name["北京飞箭商贸有限公司"]["A"] != 4:
        raise SystemExit("[FATAL] 只断一天的 A 被算进去了")
    if abs(by_name["北京德恒伟业商贸有限公司"]["score"] - 14.8) > 0.05:
        raise SystemExit("[FATAL] 德恒关注分应为 14.8，得到 %s" % by_name["北京德恒伟业商贸有限公司"]["score"])
    if "只有一天的供商" in by_name:
        raise SystemExit("[FATAL] 只断一天的供商不该上榜")
    if clean_name("北京飞箭（云仓）") != "北京飞箭":
        raise SystemExit("[FATAL] 云仓后缀没去掉")
    older, newer = {"k"}, {"k", "only_new"}
    shelf = [
        {"item_id": 1, "provider_id": 1, "name": "两天", "provider_name": "甲"},
        {"item_id": 2, "provider_id": 1, "name": "一天", "provider_name": "甲"},
    ]
    picked, _, _ = select_skus(shelf, {(1, 1)}, {(1, 1), (2, 1)})
    if [p["item_id"] for p in picked] != [1]:
        raise SystemExit("[FATAL] 连续两天才计入，选出来的是 %s" % picked)
    names = {1: "北京飞箭商贸有限公司", 2: "另一家"}
    zero_shelf = [{"item_id": 9, "provider_id": 0, "name": "自营链接", "provider_name": "0", "flow": "非流通"}]
    picked, resolved, n_fb = select_skus(zero_shelf, {(9, 1)}, {(9, 1)}, names)
    if n_fb != 1 or picked[0]["provider_id"] != 1 or resolved[0]["provider_id"] != 1:
        raise SystemExit("[FATAL] 上架供商为 0 时应改记快照供商，得到 %s" % picked)
    named = [{"item_id": 8, "provider_id": 2, "name": "已填供商", "provider_name": "另一家"}]
    picked, _, n_fb = select_skus(named, {(8, 1)}, {(8, 1)}, names)
    if picked or n_fb:
        raise SystemExit("[FATAL] 上架已填供商时不应改用快照里的另一家")
    multi = [{"item_id": 7, "provider_id": 0, "name": "两家都断", "provider_name": "0"}]
    picked, _, n_fb = select_skus(multi, {(7, 1), (7, 2)}, {(7, 1), (7, 2)}, names)
    if picked or n_fb:
        raise SystemExit("[FATAL] 多个快照供商都断时不应猜一家")
    rule_text(WEIGHT_DEFAULT)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_selfcheck_supply.xlsx")
    out = os.path.abspath(out)
    build_xlsx(rows, "2026-11-03", "2026-11-04", FLOW_CANON, WEIGHT_DEFAULT, out)
    verify_workbook(out, rows)
    os.remove(out)
    print("[self-check] 通过：等级边界、两天才计入、上架供商为 0 时改记快照供商、关注分 31 / 14.8、两页说明、卖断货色板")


# ---------- 连接（与 report_soldout.DB 相同，不改那个文件） ----------
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


class DB(object):
    def __init__(self, cfg):
        self.conn = cfg["conn"]
        self.db = cfg["db"]
        self.tunnel_ports = {}
        if self.conn.get("tunnel"):
            self.tunnel_ports = {
                "order": self.conn["tunnel"].get("order_port", 13307),
                "goods": self.conn["tunnel"].get("goods_port", 13306),
            }
        self.cli = None

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
        targets = self.conn["targets"]
        for kind in ("order", "goods"):
            tg = targets[kind]
            threading.Thread(target=bridge, args=(self._port(kind), tg["host"], tg["port"], t), daemon=True).start()
            time.sleep(0.4)
        self.cli = cli

    def q(self, kind, sql):
        import pymysql
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


def require_date(text):
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", text or ""):
        raise SystemExit("[FATAL] 日期必须是 YYYY-MM-DD")
    return text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--self-check", action="store_true")
    ap.add_argument("--config", default=None)
    ap.add_argument("--date", default=None)
    ap.add_argument("--outdir", default="./report_out")
    args = ap.parse_args()
    if args.self_check:
        self_check()
        return
    if not args.config:
        raise SystemExit("[FATAL] 需要 --config，指向本次会话上传的 ecommerce_bot.yaml")
    import yaml
    with open(args.config, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    tables = cfg["tables"]
    so = cfg["caliber"]["soldout"]
    flow_cats = list(so["flow_cats"])
    grade_th = {"流通": so["grade"]["flow"], "非流通": so["grade"]["non_flow"]}
    weights, top_n = weights_from(cfg)
    rule_text(weights)
    db = DB(cfg)
    db.connect()
    t0 = time.time()
    try:
        if args.date:
            newer = require_date(args.date)
            rows = db.q("order", "SELECT DISTINCT snapshot_date FROM %s WHERE snapshot_date IS NOT NULL AND snapshot_date <= '%s' ORDER BY snapshot_date DESC LIMIT 2" % (tables["inventory_sale"], newer))
        else:
            rows = db.q("order", "SELECT DISTINCT snapshot_date FROM %s WHERE snapshot_date IS NOT NULL ORDER BY snapshot_date DESC LIMIT 2" % tables["inventory_sale"])
        dates = [str(r[0]) for r in rows]
        if args.date and (not dates or dates[0] != args.date):
            raise SystemExit("[FATAL] %s 没有库存快照，停止。不要用别的日子代替。" % args.date)
        if len(dates) < 2:
            raise SystemExit("[FATAL] 历史不足：库存日只有 %s。连续 2 个库存日还没凑齐，停止，不要把 1 天当成 2 天出表。" % (dates or "0"))
        newer, older = dates[0], dates[1]
        print("[1/6] 库存日 %s 与 %s" % (older, newer))

        items = db.q("goods", "SELECT item_id, name, provider_id FROM %s WHERE status=%s AND sale_type=%s" % (
            tables["item_sale"], int(so["on_sale_status"]), int(so["sale_type"])))
        on_shelf = []
        seen = set()
        for item_id, name, provider_id in items:
            item_id = int(item_id)
            if item_id in seen:
                continue
            seen.add(item_id)
            on_shelf.append({
                "item_id": item_id,
                "name": name or "",
                "provider_id": int(provider_id or 0),
            })
        print("[2/6] 上架商品 %d" % len(on_shelf))

        prov = {}
        for pid, pname in db.q("goods", "SELECT provider_id, provider_name FROM %s WHERE status=1" % tables["item_provider"]):
            prov[int(pid)] = clean_name(pname)
        # goods.id 不是 item_sale.item_id（实测无关），必须按 item_id 关联；
        # 一码在 goods 里可能多行，DISTINCT 取首个；再经 item_category.id 字符串匹配成类目名
        top_cat, cat_name = {}, {}
        for item_id, cat in db.q("goods",
            "SELECT DISTINCT i.item_id, g.top_cat FROM %s i JOIN %s g ON g.item_id=i.item_id WHERE i.status=%s AND i.sale_type=%s" % (
                tables["item_sale"], tables["goods"], int(so["on_sale_status"]), int(so["sale_type"]))):
            top_cat.setdefault(int(item_id), "" if cat is None else str(cat))
        for cid, cname in db.q("goods", "SELECT id, name FROM %s" % tables["item_category"]):
            cat_name[str(cid)] = cname
        fname = {iid: cat_name.get(tc, "") for iid, tc in top_cat.items()}
        for it in on_shelf:
            it["provider_name"] = prov.get(it["provider_id"], str(it["provider_id"]))
            it["flow"] = "流通" if fname.get(it["item_id"]) in flow_cats else "非流通"

        snap_rows = db.q("order", "SELECT snapshot_date, item_id, provider_id, is_sold_out FROM %s WHERE snapshot_date IN ('%s','%s')" % (
            tables["inventory_sale"], older, newer))
        older_out, newer_out = set(), set()
        zero_rows = {}
        for snap_date, item_id, provider_id, sold in snap_rows:
            if int(provider_id or 0) == 0:
                zero_rows.setdefault(int(item_id), {})[str(snap_date)[:10]] = int(sold or 0)
            if int(sold or 0) != 1:
                continue
            key = (int(item_id), int(provider_id))
            if str(snap_date) == older:
                older_out.add(key)
            elif str(snap_date) == newer:
                newer_out.add(key)
        sold_base, on_shelf, n_fallback = select_skus(on_shelf, older_out, newer_out, prov, zero_rows)
        print("[3/6] 连续两天售罄 %d，其中上架供商为 0、改记快照供商 %d" % (len(sold_base), n_fallback))

        monthly = {}
        ids = [s["item_id"] for s in sold_base]
        year_start = newer[:4] + "-01-01"
        year_since = int(time.mktime(time.strptime(year_start, "%Y-%m-%d")))
        for i in range(0, len(ids), 400):
            chunk = ids[i:i + 400]
            if not chunk:
                continue
            ph = ",".join("'%s'" % x for x in chunk)
            sql = ("SELECT od.sku_code, SUM(od.price*od.quantity) "
                   "FROM %s od JOIN %s oh ON od.order_code=oh.order_code "
                   "WHERE od.sku_code IN (%s) AND oh.order_time>=%s AND oh.valid=1 "
                   "GROUP BY od.sku_code, DATE_FORMAT(FROM_UNIXTIME(oh.order_time),'%%Y-%%m')" % (
                       tables["order_detail"], tables["order_head"], ph, year_since))
            for item_id, amt in db.q("order", sql):
                try:
                    key = int(item_id)
                except (TypeError, ValueError):
                    key = item_id
                monthly[key] = max(monthly.get(key, 0.0), float(amt or 0))
        print("[4/6] 月销 %d 个有销售" % len(monthly))

        sold_skus = []
        for s in sold_base:
            gr = grade_of(monthly.get(s["item_id"], 0.0), s["flow"], grade_th)
            sold_skus.append(dict(s, grade=gr))
        suppliers = rank_suppliers(on_shelf, sold_skus, weights, top_n)
        print("[5/6] 上榜 %d 家" % len(suppliers))

        os.makedirs(args.outdir, exist_ok=True)
        tag = newer.replace("-", "")
        xlsx_path = os.path.join(args.outdir, "04-供货商快缺货_%s.xlsx" % tag)
        json_path = os.path.join(args.outdir, "04-供货商快缺货_%s.json" % tag)
        labels = flow_label(flow_cats)
        build_xlsx(suppliers, older, newer, labels, weights, xlsx_path)
        verify_workbook(xlsx_path, suppliers)
        slim = []
        for s in suppliers:
            one = dict((k, v) for k, v in s.items() if k != "skus")
            one["skus"] = [{"item_id": x["item_id"], "name": x["name"], "grade": x["grade"]} for x in s["skus"]]
            slim.append(one)
        payload = {
            "meta": {
                "older_date": older,
                "newer_date": newer,
                "weights": weights,
                "top_n": top_n,
                "supplier_count": len(suppliers),
                "elapsed_s": round(time.time() - t0, 2),
            },
            "suppliers": slim,
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=1)
        print("[6/6] 完成 %s" % xlsx_path)
        print("JSON:", json_path)
    finally:
        db.close()


if __name__ == "__main__":
    main()
