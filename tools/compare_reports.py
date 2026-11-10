#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报表一致性比对工具 —— 用于「新号安装结果」vs「原号产出」的逐格核对。

用法：
  # 单文件对单文件
  python3 tools/compare_reports.py --a A.xlsx --b B.xlsx

  # 整个目录对目录（按文件名自动配对）
  python3 tools/compare_reports.py --a 新号/test1 --b 本号/test1 --strict-name

  # 忽略时间戳类差异（默认已忽略）
  python3 tools/compare_reports.py --a A.xlsx --b B.xlsx --no-ignore-timestamp

判定口径：
  · sheet 名集合必须相同
  · 每个 sheet 的 max_row / max_col 必须相同
  · 逐格比对 value（数值容差 1e-9）
  · 默认忽略形如 2026-11-08T10:23:51 / 耗时 38.9s 的运行时痕迹
退出码：0 完全一致；1 存在差异；2 参数或文件问题
"""
import argparse
import os
import re
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("缺少 openpyxl，请先安装：pip install openpyxl", file=sys.stderr)
    sys.exit(2)

# 运行时痕迹：时间戳、耗时、生成时刻
TS_PATTERNS = [
    re.compile(r"\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?"),   # 2026-11-08 10:23:51
    re.compile(r"耗时\s*[:：]?\s*\d+(\.\d+)?\s*(s|秒)"),          # 耗时 38.9s
    re.compile(r"\d+(\.\d+)?\s*(s|秒)\b"),                        # 38.9s
    re.compile(r"生成于\s*[^\s|]*"),
]


def norm(v, ignore_ts):
    """归一化单元格值，便于跨实例比对。"""
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return round(float(v), 9)
    s = str(v).strip()
    if ignore_ts:
        for p in TS_PATTERNS:
            s = p.sub("<TS>", s)
    return s


def same(a, b):
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) < 1e-9
    return a == b


def cmp_workbook(pa, pb, ignore_ts, max_diffs=200):
    wa = load_workbook(pa, data_only=True)
    wb = load_workbook(pb, data_only=True)
    diffs = []
    notes = []

    names_a = set(wa.sheetnames)
    names_b = set(wb.sheetnames)
    if names_a != names_b:
        only_a = sorted(names_a - names_b)
        only_b = sorted(names_b - names_a)
        if only_a:
            diffs.append(f"[sheet] {pa} 独有：{only_a}")
        if only_b:
            diffs.append(f"[sheet] {pb} 独有：{only_b}")
        notes.append(f"sheet 名集合不一致：A={sorted(names_a)} B={sorted(names_b)}")

    for name in sorted(names_a & names_b):
        sa, sb = wa[name], wb[name]
        if (sa.max_row, sa.max_column) != (sb.max_row, sb.max_column):
            diffs.append(
                f"[shape] {name}: A={sa.max_row}R×{sa.max_column}C  "
                f"B={sb.max_row}R×{sb.max_column}C"
            )
        for r in range(1, min(sa.max_row, sb.max_row) + 1):
            for c in range(1, min(sa.max_column, sb.max_column) + 1):
                va = norm(sa.cell(r, c).value, ignore_ts)
                vb = norm(sb.cell(r, c).value, ignore_ts)
                if not same(va, vb):
                    if len(diffs) < max_diffs:
                        diffs.append(
                            f"[cell] {name}!R{r}C{c}  A={va!r}  B={vb!r}"
                        )
                    else:
                        diffs.append(f"... 差异已截断（超过 {max_diffs} 条）")
                        break
            else:
                continue
            break
    wa.close(); wb.close()
    return diffs, notes


def collect(arg):
    p = Path(arg).expanduser().resolve()
    if not p.exists():
        print(f"路径不存在：{p}", file=sys.stderr)
        sys.exit(2)
    if p.is_dir():
        return sorted([x for x in p.iterdir() if x.suffix.lower() in (".xlsx", ".xlsm")])
    return [p]


def main():
    ap = argparse.ArgumentParser(description="报表一致性比对")
    ap.add_argument("--a", required=True, help="基准（本号产出）：文件或目录")
    ap.add_argument("--b", required=True, help="待验（新号产出）：文件或目录")
    ap.add_argument("--no-ignore-timestamp", action="store_true",
                    help="不忽略时间戳/耗时类差异（默认忽略）")
    ap.add_argument("--max-diffs", type=int, default=200, help="最多列出多少条差异")
    args = ap.parse_args()

    files_a = collect(args.a)
    files_b = collect(args.b)
    if not files_a:
        print("基准侧没有 xlsx 文件", file=sys.stderr); sys.exit(2)

    pairs = []
    if len(files_a) == 1 and len(files_b) == 1 and \
            (Path(args.a).is_file() and Path(args.b).is_file()):
        pairs = [(files_a[0], files_b[0])]
    else:
        map_b = {f.name: f for f in files_b}
        for fa in files_a:
            if fa.name in map_b:
                pairs.append((fa, map_b[fa.name]))
            else:
                print(f"⚠️  待验侧缺文件：{fa.name}")
        extra = set(map_b) - {fa.name for fa in files_a}
        if extra:
            print(f"⚠️  待验侧多出文件：{sorted(extra)}")

    if not pairs:
        print("没有可配对的文件", file=sys.stderr); sys.exit(2)

    total_bad = 0
    for fa, fb in pairs:
        diffs, notes = cmp_workbook(fa, fb, not args.no_ignore_timestamp, args.max_diffs)
        print(f"\n=== {fa.name} ===")
        if not diffs:
            print("  ✅ 完全一致")
        else:
            total_bad += 1
            print(f"  ❌ 发现 {len(diffs)} 处差异")
            for d in diffs[:50]:
                print("   ", d)
            if len(diffs) > 50:
                print(f"    …… 另有 {len(diffs) - 50} 条未显示")
        for n in notes:
            print(f"  ℹ️  {n}")

    print("\n" + "=" * 50)
    if total_bad == 0:
        print(f"结论：{len(pairs)} 个文件全部一致 ✅")
        return 0
    print(f"结论：{total_bad}/{len(pairs)} 个文件存在差异 ❌")
    return 1


if __name__ == "__main__":
    sys.exit(main())
