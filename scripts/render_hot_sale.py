#!/usr/bin/env python3
"""② 爆卖双榜 XLSX 出表器（样式锁死，禁止现场自写）

用法：
  python3 scripts/render_hot_sale.py --json <爆卖JSON> --out <输出XLSX>

输入 JSON：{"meta":{"day","window","counts"},"size":[...],"speed":[...],"new":[...]}
行字段：tier item_id name provider_name fname qty_7d avg_3d acc rise_days stock_grade cover_days

样式基线 = ① 卖断货 / ④ 供货缺货 已上线契约。本文件是唯一允许的出表入口。

版式（三 sheet）：
  「规模爆品」「速度爆品」明细页：r1 标题 / r2 副标题 / r3 表头（深底白字）/ r4 起数据。
    明细页只留数据和列，**不放规则块**——老板打开直接看到商品，不用滚 16 行。
  「规则总览」：两榜的入选条件 + 四档看板 + 通用指标与备注规则，集中一页。
文字控制在合并宽度内，避免渲染截断。
"""
import argparse
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

FONT = "Microsoft YaHei"
C_TITLE = "111827"
C_SUB = "6B7280"
C_HEAD = "1F2937"

TIERS = [
    ("DC2626", "FEE2E2", "B91C1C"),
    ("C2410C", "FFEDD5", "C2410C"),
    ("A16207", "FEF9C3", "A16207"),
    ("4B5563", "F3F4F6", "4B5563"),
]

SIZE_TIERS = ["爆款", "旺销", "畅销", "起量"]
SPEED_TIERS = ["黑马", "猛涨", "起势", "新品"]

SIZE_COLS = ["档位", "商城商品码", "商品名称", "供商", "大分类",
             "近7日销量", "近3日日均", "加速比", "上涨天数", "库存价值档", "可售天数",
             "备注（异常原因）"]
SIZE_W = [10, 12, 36, 26, 10, 11, 10, 10, 10, 11, 10, 46]

SPEED_COLS = ["档位", "商城商品码", "商品名称", "供商", "大分类",
              "近7日销量", "加速比", "上涨天数", "近3日日均", "库存价值档", "可售天数",
              "备注（异常原因）"]
SPEED_W = [10, 12, 36, 26, 10, 11, 11, 10, 10, 11, 10, 46]

SIZE_INTRO = (
    "近 7 日销量 ≥50 件　且　加速比 ≥1.5　且　上涨天数 ≥3 天，三条同时满足才进榜；"
    "当日售罄的品不进榜（与卖断货同判据）。"
)
SIZE_METRICS = [
    "加速比 = 近 3 日日均 ÷ 前 4 日日均；前 4 日为 0 且有量 → 记「新增」，不算无穷大",
    "上涨天数 = 只数有货且销量递增的天数（断货日的 0 不计）；可售天数 = 当前库存 ÷ 近 3 日日均",
    "库存价值档 = 与卖断货同口径：流通品月销额 >1 万、非流通 >2 千为 A；1000 整为 B、500 整为 C",
    "备注列 = 可售 >1000 天「云仓 · 天数」（云仓=有货量大/无货即断）、731~1000 天「库存偏高」、峰值超其余 6 日均值 5 倍「含单日尖峰」、供商空「没有供商」",
]
SIZE_BOARD = [
    ("爆款", "≥500 件", "单周最大体量，优先保货"),
    ("旺销", "200~500 件", "稳定走量，别让断"),
    ("畅销", "100~200 件", "持续放量，盯一周趋势"),
    ("起量", "50~100 件", "刚起量，观察一周"),
]

SPEED_INTRO = (
    "近 7 日 ≥10 件（不设上限）　+　近 3 日日均 ≥5 件　+　加速比 ≥3 倍　+　上涨天数 ≥3 天　+　"
    "剔单日尖峰　+　后 3 天无 0 断点　+　非「缺货→补货」恢复型，全部满足才进榜。"
)
SPEED_METRICS = [
    "加速比 = 近 3 日日均 ÷ 前 4 日日均；前 4 日为 0 且有量 → 记「新增」，新品档即此类",
    "上涨天数 = 只数有货且销量递增的天数（断货日的 0 不计）；可售天数 = 当前库存 ÷ 近 3 日日均",
    "库存价值档 = 与卖断货同口径：流通品月销额 >1 万、非流通 >2 千为 A；1000 整为 B、500 整为 C",
    "备注列 = 可售 >1000 天「云仓 · 天数」（云仓=有货量大/无货即断）、731~1000 天「库存偏高」、峰值超其余 6 日均值 5 倍「含单日尖峰」、供商空「没有供商」",
]
SPEED_BOARD = [
    ("黑马", "≥10 倍", "异常陡峭，最快爆"),
    ("猛涨", "5~10 倍", "快速拉升，重点补货"),
    ("起势", "3~5 倍", "开始抬头，先盯住"),
    ("新品", "前4日=0", "刚起量的新面孔"),
]

HEAD_ROW = 3
DATA_ROW = 4


# 统一网格：四张表共用同一组细线，避免各自风格打架
_SIDE = Side(style="thin", color="E5E7EB")
GRID = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)


def paint_grid(ws, r1, r2, c1, c2):
    """给 r1..r2 行、c1..c2 列刷统一细网格（不碰填充与字体）。"""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = GRID


def _tier_index(labels, tier):
    try:
        return labels.index(tier)
    except ValueError:
        return None


def build_sheet(ws, title, subtitle, tier_labels, cols, widths, rows,
                key_col, key_size):
    n = len(cols)
    last_col = get_column_letter(n)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    c = ws.cell(1, 1, title)
    c.font = Font(name=FONT, size=16, bold=True, color=C_TITLE)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    c = ws.cell(2, 1, subtitle)
    c.font = Font(name=FONT, size=10, color=C_SUB)
    c.alignment = Alignment(horizontal="left", vertical="center")

    for i, name in enumerate(cols):
        cell = ws.cell(HEAD_ROW, i + 1, name)
        cell.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=C_HEAD)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[HEAD_ROW].height = 22

    labels = tier_labels
    for ri, row in enumerate(rows):
        r = DATA_ROW + ri
        for ci, val in enumerate(row):
            cell = ws.cell(r, ci + 1, val)
            last_cell = ci == len(row) - 1
            cell.font = Font(name=FONT, size=10, color=C_SUB if last_cell else C_TITLE)
            cell.alignment = Alignment(horizontal="left" if last_cell else "center",
                                       vertical="center", wrap_text=last_cell)
        ti = _tier_index(labels, row[0])
        if ti is not None:
            _dark, light, tc = TIERS[ti]
            cell = ws.cell(r, 1)
            cell.font = Font(name=FONT, size=10, bold=True, color=tc)
            cell.fill = PatternFill("solid", fgColor=light)
        cell = ws.cell(r, key_col)
        cell.font = Font(name=FONT, size=key_size, bold=True, color=C_TITLE)

    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    last = HEAD_ROW + len(rows)
    paint_grid(ws, HEAD_ROW, last, 1, n)      # 表头 + 数据统一细网格
    ws.auto_filter.ref = f"A{HEAD_ROW}:{last_col}{last}"
    ws.freeze_panes = f"A{DATA_ROW}"
    ws.sheet_view.showGridLines = False
    # 打印横向 + 缩放到一页宽，避免规则行与右侧列在纸质版被裁
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def build_overview(ws, day, win, size_total, speed_total, board_s, board_p):
    """规则总览页：两榜入选条件 + 四档看板 + 通用口径，独立一页，明细页不再放规则。"""
    n = 4

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    c = ws.cell(1, 1, "爆卖双榜 · 规则总览")
    c.font = Font(name=FONT, size=16, bold=True, color=C_TITLE)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    c = ws.cell(2, 1, f"数据日期 {day}｜窗口 {win[0]} ~ {win[1]}｜规模 {size_total} 个、速度 {speed_total} 个")
    c.font = Font(name=FONT, size=10, color=C_SUB)
    c.alignment = Alignment(horizontal="left", vertical="center")

    def block(start, heading, intro, board):
        c = ws.cell(start, 1, heading)
        c.font = Font(name=FONT, size=12, bold=True, color=C_TITLE)
        c.alignment = Alignment(horizontal="left", vertical="center")

        r = start + 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
        c = ws.cell(r, 1, intro)
        c.font = Font(name=FONT, size=10, color=C_TITLE)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 48

        r = start + 2
        ws.row_dimensions[r].height = 40
        for i, (label, cnt, thresh, _desc) in enumerate(board):
            cc = i + 1
            _dark, light, tc = TIERS[i]
            cell = ws.cell(r, cc, cnt)
            cell.font = Font(name=FONT, size=22, bold=True, color=tc)
            cell.fill = PatternFill("solid", fgColor=light)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell = ws.cell(r + 1, cc, label)
            cell.font = Font(name=FONT, size=11, bold=True, color=tc)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell = ws.cell(r + 2, cc, thresh)
            cell.font = Font(name=FONT, size=11, color=C_TITLE)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        r = start + 5
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
        c = ws.cell(r, 1, "   |   ".join(f"{b[0]} {b[3]}" for b in board))
        c.font = Font(name=FONT, size=10, color=C_SUB)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 24
        return start + 7

    nxt = block(4, "一、规模爆品（看体量：谁卖得多）", SIZE_INTRO, board_s)
    nxt = block(nxt, "二、速度爆品（看加速度：谁涨得猛）", SPEED_INTRO, board_p)

    c = ws.cell(nxt, 1, "三、通用口径（两榜共用）")
    c.font = Font(name=FONT, size=12, bold=True, color=C_TITLE)
    c.alignment = Alignment(horizontal="left", vertical="center")
    for i, text in enumerate(SIZE_METRICS[:4]):
        r = nxt + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
        c = ws.cell(r, 1, text)
        c.font = Font(name=FONT, size=10, color=C_TITLE)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 32

    for i in range(1, n + 1):
        ws.column_dimensions[get_column_letter(i)].width = 50
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def _human_cover(v):
    if not v:
        return "—"
    return f"{v/10000:.0f}万天" if v >= 10000 else f"{v:.0f}天"


def anomaly_note(x):
    """异常原因，最多两条；判据写在表内 r15，改规则必须同步。"""
    notes = []
    cov = x.get("cover_days") or 0
    ser = x.get("series") or []
    spike = None
    if len(ser) == 7 and max(ser) >= 10:
        peak = max(range(7), key=lambda k: ser[k])
        rest = [v for k, v in enumerate(ser) if k != peak]
        base = sum(rest) / len(rest) if rest else 0
        if base > 0 and ser[peak] / base > 5:
            spike = ser[peak]
    if cov > 1000:
        notes.append(f"云仓 · {_human_cover(cov)}")
    if spike:
        notes.append(f"含单日尖峰 {spike:.0f}件")
    if 730 < cov <= 1000:
        notes.append(f"库存偏高 {_human_cover(cov)}")
    if not x.get("provider_name"):
        notes.append("没有供商")
    return "；".join(notes[:2])


def fmt_size_row(x):
    return [x["tier"], x["item_id"], x["name"], x.get("provider_name") or "—",
            x.get("fname") or "—", int(x["qty_7d"]), x["avg_3d"], x["acc"],
            x["rise_days"], x["stock_grade"], x["cover_days"], anomaly_note(x)]


def fmt_speed_row(x):
    acc = x["acc"] if x.get("acc") is not None else "新增"
    return [x["tier"], x["item_id"], x["name"], x.get("provider_name") or "—",
            x.get("fname") or "—", int(x["qty_7d"]), acc, x["rise_days"],
            x["avg_3d"], x["stock_grade"], x["cover_days"], anomaly_note(x)]


def render(payload, out_path):
    meta = payload.get("meta", {})
    day = meta.get("day", "")
    win = meta.get("window") or ["", ""]
    size = payload.get("size", [])
    speed = payload.get("speed", [])
    new = payload.get("new", [])

    wb = Workbook()

    ws0 = wb.active
    ws0.title = "规则总览"

    ws1 = wb.create_sheet("规模爆品")
    c1 = {g: sum(1 for x in size if x["tier"] == g) for g in SIZE_TIERS}
    board1 = [(label, c1[label], thresh, desc)
              for (label, thresh, desc) in SIZE_BOARD]
    build_sheet(
        ws1,
        "爆卖规模榜 · 规模爆品",
        f"数据日期 {day}｜窗口 {win[0]} ~ {win[1]}｜共 {len(size)} 个｜按档位、销量降序",
        SIZE_TIERS,
        SIZE_COLS, SIZE_W, [fmt_size_row(x) for x in size], 6, 11,
    )
    for r in range(DATA_ROW, DATA_ROW + len(size)):
        ws1.cell(r, 8).number_format = '0.0"x"'
        ws1.cell(r, 11).number_format = "0.0"

    ws2 = wb.create_sheet("速度爆品")
    c2 = {g: sum(1 for x in speed if x["tier"] == g) for g in SPEED_TIERS[:3]}
    c2["新品"] = len(new)
    board2 = [(label, c2[label], thresh, desc)
              for (label, thresh, desc) in SPEED_BOARD]
    rows2 = [fmt_speed_row(x) for x in speed] + [fmt_speed_row(x) for x in new]
    build_sheet(
        ws2,
        "爆卖速度榜 · 速度爆品",
        f"数据日期 {day}｜窗口 {win[0]} ~ {win[1]}｜共 {len(rows2)} 个"
        f"（黑马/猛涨/起势 {len(speed)} + 新品 {len(new)}）｜按档位、加速比降序",
        SPEED_TIERS,
        SPEED_COLS, SPEED_W, rows2, 7, 12,
    )
    for r in range(DATA_ROW, DATA_ROW + len(rows2)):
        ws2.cell(r, 11).number_format = "0.0"

    build_overview(ws0, day, win, len(size), len(rows2), board1, board2)

    out_dir = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(out_dir, exist_ok=True)
    wb.save(out_path)
    return {"size": len(size), "speed": len(rows2), "out": out_path,
            "size_tier": c1, "speed_tier": c2}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    with open(a.json, encoding="utf-8") as f:
        payload = json.load(f)
    res = render(payload, a.out)
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    sys.exit(main())
