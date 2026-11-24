#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
月销本地缓存只读复用（②③④ 共用 ① 卖断货落盘的月销缓存）。

缓存文件：~/.qpwy_bot/soldout_monthly_cache.json（由 scripts/report_soldout.py 写入并每日增量更新）
结构：{"year_since": <unix 秒>, "last_ym": "YYYY-MM", "items": {sku: {"YYYY-MM": 月销额}}, "updated": ...}

口径（与 soldout.yaml 第 4 节一致）：
  月销额 = SUM(price * quantity)，valid=1，自然年各月；等级判据取自然年内最高一个月。

本模块只**读**缓存，不写、不重建；缓存缺失/损坏/年份不符时返回 None，
调用方必须回退到原有现查 SQL（兜底路径保持不变）。
"""
import json
import os

MONTHLY_CACHE = os.path.expanduser("~/.qpwy_bot/soldout_monthly_cache.json")


def load_monthly_items(year=None):
    """读取月销缓存 items。

    :param year: 目标自然年（如 2026）；传入则校验缓存月份属于该年，不符返回 None。
    :return: dict {sku(str): {"YYYY-MM": amt(float)}}；不可用时返回 None。
    """
    try:
        if not os.path.exists(MONTHLY_CACHE):
            return None
        with open(MONTHLY_CACHE, encoding="utf-8") as f:
            c = json.load(f)
        items = c.get("items")
        if not isinstance(items, dict) or not items:
            return None
        if year is not None:
            years = {m.split("-")[0] for d in items.values() for m in d.keys()}
            if str(year) not in years:
                return None
        return items
    except Exception:
        return None


def max_month_amt(items, sku):
    """取某商品在缓存中的自然年内最高月销额；无记录返回 0.0。"""
    if not items:
        return 0.0
    d = items.get(str(sku))
    if not d:
        d = items.get(sku)
    if not d:
        return 0.0
    try:
        return max(float(v or 0) for v in d.values())
    except Exception:
        return 0.0
