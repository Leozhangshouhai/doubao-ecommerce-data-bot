#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
③ 热品预判缺货 出表器（JSON → XLSX）。

样式与 ①④② 同一套：微软雅黑、标题 `111827`、表头 `1F2937` 白字、
四档沿用共享色板（红/橙/黄/灰，按序位取色，不新增颜色）。
结构照 ② 明细页：r1 标题 / r2 副标 / r3 表头 / r4 起数据，auto_filter A3:L，freeze A4。

用法：
  python3 render_predict.py --json <qpwy_热品预判缺货_YYYYMMDD.json> --out <xlsx>
"""
import argparse
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

FONT = "Microsoft YaHei"
C_TITLE, C_SUB, C_HEAD = "111827", "6B7280", "1F2937"
# 深底 / 浅底 / 字色，按序位：P0 红 → P3 灰
TIERS = [("DC2626", "FEE2E2", "B91C1C"),
         ("C2410C", "FFEDD5", "C2410C"),
         ("A16207", "FEF9C3", "A16207"),
         ("4B5563", "F3F4F6", "4B5563")]
LEVELS = ["P0", "P1", "P2", "P3"]
LEVEL_NAME = {"P0": "P0 今天下单", "P1": "P1 明天下单", "P2": "P2 本周安排", "P3": "P3 观察"}
LEVEL_DES = {"P0": "可售 ≤1 天", "P1": "≤2 天", "P2": "≤5 天", "P3": "≤15 天"}
LEVEL_WHAT = {"P0": "今天就断，T+1 才接得上", "P1": "明天必须下单", "P2": "本周内安排", "P3": "货还够，先盯着"}

COLS = ["优先级", "商城商品码", "商品名称", "供商", "大分类",
        "近7日销量", "日均销量", "当前库存", "可售天数", "预估断货日", "建议下单日", "是否爆卖榜"]
WIDTHS = [9, 12, 30, 24, 9, 10, 9, 9, 9, 12, 12, 10]
HEAD_ROW, DATA_ROW = 3, 4


def yh(size, bold=False, color=None):
    return Font(name=FONT, size=size, bold=bold, color=color)


def title_row(ws, t, sub, last_col):
    ws.cell(1, 1, t).font = yh(16, True, C_TITLE)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.row_dimensions[1].height = 28
    ws.cell(2, 1, sub).font = yh(10, False, C_SUB)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)


def hdr_row(ws, headers, row):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row, j, h)
        c.fill = PatternFill("solid", fgColor=C_HEAD)
        c.font = yh(11, True, "FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 22


def build_overview(ws, d):
    cnt = d.get("counts", {})
    pool = d.get("pool", 0)
    total = len(d.get("rows", []))
    th = d.get("thresholds", {})
    title_row(ws, "热品预判缺货 · 一眼看懂",
              "数据日期 %s｜热品池 %d｜预警 %d｜口径 v1.0（热品池 ∩ 可售 ≤%g 天）"
              % (d["day"], pool, total, th.get("max_cover_days", 15)), 4)

    ws.cell(4, 1, "按紧急度数量").font = yh(12, True)
    for i, lv in enumerate(LEVELS):
        col = i + 1
        deep, light, tone = TIERS[i]
        num = ws.cell(5, col, cnt.get(lv, 0))
        num.fill = PatternFill("solid", fgColor=light)
        num.font = yh(20, True, tone)
        num.alignment = Alignment(horizontal="center", vertical="center")
        lab = ws.cell(6, col, LEVEL_NAME[lv])
        lab.font = yh(11, True, tone)
        lab.alignment = Alignment(horizontal="center", vertical="center")
        des = ws.cell(7, col, LEVEL_DES[lv])
        des.font = yh(10, False, C_SUB)
        des.alignment = Alignment(horizontal="center", vertical="center")
        what = ws.cell(8, col, LEVEL_WHAT[lv])
        what.font = yh(9, False, C_SUB)
        what.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[5].height = 40
    ws.row_dimensions[8].height = 30

    ws.cell(10, 1, "怎么用（建议）").font = yh(12, True)
    notes = [
        "1. 本表是提前预警，不是已断货名单：已断的看 ① 卖断货与 ④ 供货缺货，这里只收「有货但撑不住」的品。",
        "2. 优先级只看可售天数：P0 ≤1 天（今天下单，T+1 到货才接得上）→ P1 ≤2 天 → P2 ≤5 天 → P3 ≤15 天。",
        "3. 日均取 max(近3日均, 近7日日均)，宁可早报：日报当天突然放量时不会把紧急度算轻。",
        "4. 「是否爆卖榜」= 同时命中 ② 的品，也就是「又涨又快断」，采购应最先处理这一批。",
        "5. 建议下单日 = 预估断货日 −1；留一天给供商备货与物流，不要等到断货当天再下。",
    ]
    for i, nt in enumerate(notes):
        r = 11 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(r, 1, nt).font = yh(10)
        ws.cell(r, 1).alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 36

    ws.cell(17, 1, "紧急度 TOP 供商（P0+P1）").font = yh(12, True)
    hdr_row(ws, ["供商", "P0+P1 数量", "其中 P0", "其中 P1"], 18)
    ab = {}
    for r in d.get("rows", []):
        if r["level"] in ("P0", "P1"):
            p = r["provider_name"] or "没有供商"
            ab.setdefault(p, [0, 0, 0])
            ab[p][0] += 1
            ab[p][1 if r["level"] == "P0" else 2] += 1
    top = sorted(ab.items(), key=lambda x: (-x[1][0], -x[1][1]))[:8]
    for i, (nm, (t, na, nb)) in enumerate(top):
        r = 19 + i
        c0 = ws.cell(r, 1, nm)
        c0.font = yh(10)
        c0.alignment = Alignment(vertical="center", wrap_text=True)
        for j, v in enumerate([t, na, nb], 2):
            c = ws.cell(r, j, v)
            c.font = yh(10)
            c.alignment = Alignment(horizontal="center", vertical="center")
    for j in range(1, 5):
        ws.column_dimensions[chr(64 + j)].width = 26 if j == 1 else 16
    ws.print_area = "A1:D27"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def build_sheet(ws, d):
    rows = d.get("rows", [])
    title_row(ws, "热品预判缺货 · 预判清单",
              "数据日期 %s｜窗口 %s~%s｜共 %d 个（P0 %d / P1 %d / P2 %d / P3 %d）｜按可售天数升序"
              % (d["day"], d["win"][0], d["win"][-1], len(rows),
                 d["counts"].get("P0", 0), d["counts"].get("P1", 0),
                 d["counts"].get("P2", 0), d["counts"].get("P3", 0)), 12)
    hdr_row(ws, COLS, HEAD_ROW)

    for i, r in enumerate(rows):
        row = DATA_ROW + i
        idx = LEVELS.index(r["level"])
        deep, light, tone = TIERS[idx]
        a = ws.cell(row, 1, r["level"])
        a.fill = PatternFill("solid", fgColor=deep)
        a.font = yh(10, True, "FFFFFF")
        a.alignment = Alignment(horizontal="center", vertical="center")
        vals = [str(r["item_id"]), r["name"], r["provider_name"] or "没有供商", r["flow"] + "品",
                int(r["q7"]), r["daily"], int(r["stock"]), r["cover"],
                r["est_out"], r["order_by"], "是" if r["in_hot"] else ""]
        for j, v in enumerate(vals, 2):
            c = ws.cell(row, j, v)
            c.font = yh(10)
            c.alignment = Alignment(horizontal="left" if j in (3, 4) else "center",
                                    vertical="center",
                                    shrink_to_fit=(j in (3, 4)))   # 单行+超长自动缩，避免溢出压到相邻列
        ws.row_dimensions[row].height = 16
        ws.cell(row, 9).number_format = "0.0"
        ws.cell(row, 7).number_format = "0.0"
        if r["in_hot"]:
            ws.cell(row, 12).font = yh(10, True, "B91C1C")

    for j, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[chr(64 + j)].width = w
    last = DATA_ROW + len(rows) - 1
    if rows:
        ws.auto_filter.ref = "A%d:L%d" % (HEAD_ROW, last)
        ws.freeze_panes = "A%d" % DATA_ROW
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    d = json.load(open(a.json, encoding="utf-8"))
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "一眼看懂"
    ws2 = wb.create_sheet("预判清单")
    build_overview(ws1, d)
    build_sheet(ws2, d)
    wb.save(a.out)
    print("已出表: %s" % a.out)
    print("  sheet: %s｜预警 %d 条｜档位 %s" % (wb.sheetnames, len(d["rows"]), d["counts"]))


if __name__ == "__main__":
    main()
