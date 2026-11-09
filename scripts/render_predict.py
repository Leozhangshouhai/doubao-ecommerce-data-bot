#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
③ 热品预判缺货 出表器（JSON → XLSX）。

样式与 ①④② 同一套：微软雅黑、标题 `111827`、表头 `1F2937` 白字、
四档沿用共享色板（红/橙/黄/灰，按序位取色，不新增颜色）。
结构照 ② 明细页：r1 标题 / r2 副标 / r3 表头 / r4 起数据，auto_filter A3:N（14 列），freeze A4。

用法：
  python3 render_predict.py --json <qpwy_热品预判缺货_YYYYMMDD.json> --out <xlsx>
"""
import argparse
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT = "Microsoft YaHei"
C_TITLE, C_SUB, C_HEAD = "111827", "6B7280", "1F2937"
# 深底 / 浅底 / 字色，按序位：P0 红 → P3 灰
TIERS = [("DC2626", "FEE2E2", "B91C1C"),
         ("C2410C", "FFEDD5", "C2410C"),
         ("A16207", "FEF9C3", "A16207"),
         ("4B5563", "F3F4F6", "4B5563")]
LEVELS = ["P0", "P1", "P2", "P3"]
LEVEL_NAME = {"P0": "P0 立刻对接", "P1": "P1 本周下单", "P2": "P2 排期跟进", "P3": "P3 观察"}
LEVEL_DES = {"P0": "可售 ≤2 天", "P1": "3~5 天", "P2": "5~10 天", "P3": "10~20 天"}
LEVEL_WHAT = {"P0": "2 天内卖空", "P1": "5 天内断货", "P2": "10 天内断货", "P3": "20 天内断货"}

COLS = ["优先级", "等级", "商城商品码", "商品名称", "供商", "大分类",
        "近7日销量", "日均销量", "当前库存", "可售天数",
        "预计断货区间", "建议下单日", "3天缺口", "是否爆卖榜"]
WIDTHS = [9, 6, 12, 30, 22, 9, 10, 9, 9, 9, 19, 11, 10, 10]
HEAD_ROW, DATA_ROW = 3, 4


# 统一网格：四张表共用同一组细线，避免各自风格打架
_SIDE = Side(style="thin", color="E5E7EB")
GRID = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)


def paint_grid(ws, r1, r2, c1, c2):
    """给 r1..r2 行、c1..c2 列刷统一细网格（不碰填充与字体）。"""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = GRID


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
        c.font = yh(10, True, "FFFFFF")          # 11→10：筛选箭头挤占列宽，11 号会把「优先级」裁成「优」
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28           # 允许表头折两行，长列名一字不少


def build_overview(ws, d):
    cnt = d.get("counts", {})
    pool = d.get("pool", 0)
    total = len(d.get("rows", []))
    th = d.get("thresholds", {})
    title_row(ws, "热品预判缺货 · 一眼看懂",
              "数据日期 %s｜热品池 %d｜预警 %d｜口径 v1.1（热品池 ∩ 可售 ≤%g 天）"
              % (d["day"], pool, total, th.get("max_cover_days", 20)), 4)

    # 覆盖声明（两行放得下，行高 30）—— 这句必须写在脸上，避免把「只覆盖 2 类供商」误读成全站
    ws.merge_cells("A3:D3")
    cov = ws.cell(3, 1, "覆盖范围：仅库存数据可信的 2 类供商（自营 209 + 百世鼎盛 426）。"
                       "其余 87 家为云仓占位库存（写着 10000），算不出可售天数，其风险见 ① 卖断货 / ④ 供货缺货。")
    cov.font = yh(9, True, "C2410C")
    cov.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 30

    ws.cell(5, 1, "按紧急度数量").font = yh(12, True)
    for i, lv in enumerate(LEVELS):
        col = i + 1
        deep, light, tone = TIERS[i]
        num = ws.cell(6, col, cnt.get(lv, 0))
        num.fill = PatternFill("solid", fgColor=light)
        num.font = yh(20, True, tone)
        num.alignment = Alignment(horizontal="center", vertical="center")
        lab = ws.cell(7, col, LEVEL_NAME[lv])
        lab.font = yh(11, True, tone)
        lab.alignment = Alignment(horizontal="center", vertical="center")
        des = ws.cell(8, col, LEVEL_DES[lv])
        des.font = yh(10, False, C_SUB)
        des.alignment = Alignment(horizontal="center", vertical="center")
        what = ws.cell(9, col, LEVEL_WHAT[lv])
        what.font = yh(9, False, C_SUB)
        what.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[6].height = 40
    ws.row_dimensions[9].height = 30

    ws.cell(11, 1, "怎么用（建议）").font = yh(12, True)
    notes = [
        "1. 只收「有货但撑不住」的品；已断的看 ① 卖断货与 ④ 供货缺货。",
        "2. 档位按可售天数：P0≤2 / P1 3~5 / P2 5~10 / P3 10~20 天。",
        "3. 日均已剔除「塌陷日」，避免断货日被推后。",
        "4. 断货日给区间、每天滚动，不是承诺。",
        "5. 3天缺口 = 日均×3 − 库存，够撑 3 天显示 0。",
        "6. 等级 A/B/C/D 与 ① 同口径；同档排序：缺口→等级→日均。",
        "7. 自营页内部补货/调拨；热品页向百世鼎盛下单。",
    ]
    for i, nt in enumerate(notes):
        r = 12 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(r, 1, nt).font = yh(10)
        ws.cell(r, 1).alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 22

    ws.cell(20, 1, "清单构成（按供商）").font = yh(12, True)
    hdr_row(ws, ["供商", "全部预警", "其中 P0+P1", "其中 P0"], 21)
    ab = {}
    for r in d.get("rows", []):
        p = r["provider_name"] or "没有供商"
        a3 = ab.setdefault(p, [0, 0, 0])
        a3[0] += 1
        if r["level"] in ("P0", "P1"):
            a3[1] += 1
            if r["level"] == "P0":
                a3[2] += 1
    top = sorted(ab.items(), key=lambda x: (-x[1][0], -x[1][2]))[:8]
    for i, (nm, (allc, ab12, na)) in enumerate(top):
        r = 22 + i
        c0 = ws.cell(r, 1, nm)
        c0.font = yh(10)
        c0.alignment = Alignment(vertical="center", wrap_text=True)
        for j, v in enumerate([allc, ab12, na], 2):
            c = ws.cell(r, j, v)
            c.font = yh(10)
            c.alignment = Alignment(horizontal="center", vertical="center")
    paint_grid(ws, 21, 21 + len(top), 1, 4)          # 清单构成小表
    for j in range(1, 5):
        ws.column_dimensions[chr(64 + j)].width = 26 if j == 1 else 16
    ws.print_area = "A1:D29"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def build_sheet(ws, d, rows=None, title=None, sub=None):
    rows = d.get("rows", []) if rows is None else rows
    c = {k: sum(1 for r in rows if r["level"] == k) for k in LEVELS}
    title_row(ws, title or "热品预判缺货 · 预判清单",
              sub or ("数据日期 %s｜窗口 %s~%s｜共 %d 个（P0 %d / P1 %d / P2 %d / P3 %d）｜按档位 + 缺口/日均降序"
                      % (d["day"], d["win"][0], d["win"][-1], len(rows),
                         c["P0"], c["P1"], c["P2"], c["P3"])), 12)
    hdr_row(ws, COLS, HEAD_ROW)

    for i, r in enumerate(rows):
        row = DATA_ROW + i
        idx = LEVELS.index(r["level"])
        deep, light, tone = TIERS[idx]
        a = ws.cell(row, 1, r["level"])
        a.fill = PatternFill("solid", fgColor=deep)
        a.font = yh(10, True, "FFFFFF")
        a.alignment = Alignment(horizontal="center", vertical="center")
        e1 = r.get("est_out_min") or r["est_out"]
        e2 = r.get("est_out_max") or r["est_out"]
        est = e1 if e1 == e2 else "%s~%s" % (e1, str(e2)[5:])
        vals = [r.get("grade", "D"), str(r["item_id"]), r["name"], r["provider_name"] or "没有供商",
                r["flow"] + "品", int(r["q7"]), r["daily"], int(r["stock"]), r["cover"],
                est, r["order_by"], int(r.get("gap3", 0)), "是" if r["in_hot"] else ""]
        for j, v in enumerate(vals, 2):
            c = ws.cell(row, j, v)
            if j == 2:                      # 等级列：A 红 / B 橙 / C 深灰 / D 浅灰
                tone = {"A": ("DC2626", True), "B": ("C2410C", True),
                        "C": ("4B5563", False), "D": ("9CA3AF", False)}.get(str(v), ("6B7280", False))
                c.font = yh(10, tone[1], tone[0])
            else:
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
        paint_grid(ws, HEAD_ROW, last, 1, len(COLS))     # 表头 + 数据统一细网格
        ws.auto_filter.ref = "A%d:N%d" % (HEAD_ROW, last)
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
    build_overview(ws1, d)

    rows = d.get("rows", [])
    own = [r for r in rows if not r.get("provider_id")]          # 自营：没有供商
    hot = [r for r in rows if r.get("provider_id")]              # 其余真实库存供商（本期=百世鼎盛）

    def _sub(label):
        c = {k: sum(1 for r in x if r["level"] == k) for k in LEVELS}
        return ("数据日期 %s｜窗口 %s~%s｜%s 共 %d 个（P0 %d / P1 %d / P2 %d / P3 %d）｜"
                "按档位 + 缺口 + 等级 + 日均降序"
                % (d["day"], d["win"][0], d["win"][-1], label, len(x), c["P0"], c["P1"], c["P2"], c["P3"]))

    for name, label, ttl in (("自营补货清单", "自营（没有供商）", "热品预判缺货 · 自营补货清单"),
                             ("热品清单", "百世鼎盛等真实库存供商", "热品预判缺货 · 热品清单")):
        x = own if name == "自营补货清单" else hot
        ws = wb.create_sheet(name)
        build_sheet(ws, d, x, ttl, _sub(label))

    wb.save(a.out)
    print("已出表: %s" % a.out)
    print("  sheet: %s｜自营 %d｜其他 %d｜合计 %d" % (wb.sheetnames, len(own), len(hot), len(rows)))


if __name__ == "__main__":
    main()
