#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SKILL 触发词冲突检查

用途：新建 SKILL 前跑一次，确认它的 description 没有覆盖已有模块的触发词域。
只读：只扫描目录，不修改任何文件。

用法：
  python3 tools/check_skill_conflict.py
  python3 tools/check_skill_conflict.py --root ~/.doubao/agent_mode/workspace/.user_skills
  python3 tools/check_skill_conflict.py --json      # 机器可读输出

退出码：0 = 无冲突；1 = 有冲突；2 = 扫描失败
"""
import argparse
import json
import os
import re
import sys

# 只过滤虚词，业务词一律不过滤（"报表""数据"这类要不要算冲突，交给人判断）
STOP = {
    "什么", "怎么", "可以", "需要", "使用", "支持", "包括", "以及", "或者",
    "并且", "如果", "然后", "这个", "那个", "一个", "进行", "通过", "相关",
    "以及", "其中", "对于", "关于", "能够", "是否", "已经", "所有", "任何",
    "时候", "情况", "问题", "功能", "能力", "场景", "任务", "内容", "方式",
    # 描述框架词，不是业务触发词，参与比对只会产生噪音
    "触发词", "触发场景", "触发条件", "时使用", "时调用", "用户提到", "当用户",
    "本技能", "本skill", "skill", "工具", "插件",
}


def read_frontmatter(path):
    """返回 (name, description)；解析失败返回 (None, None)"""
    try:
        txt = open(path, encoding="utf-8").read()
    except Exception:
        return None, None
    m = re.match(r"^---\s*\n(.*?)\n---", txt, re.S)
    if not m:
        return None, None
    fm = m.group(1)
    name_m = re.search(r"^\s*name:\s*(.+)$", fm, re.M)
    name = name_m.group(1).strip().strip("\"'") if name_m else None
    # description 可能跨行：取到下一个顶层 key 之前
    dm = re.search(r"^\s*description:\s*(.+?)(?=\n\s*[A-Za-z_][A-Za-z0-9_-]*:\s|\Z)", fm, re.S | re.M)
    desc = " ".join(dm.group(1).split()) if dm else None
    return name, desc


def candidates(desc):
    """从 description 提候选触发词：优先「」引用，其次按标点切分"""
    out = set()
    for w in re.findall(r"「([^」]{1,20})」", desc):
        out.add(w.strip())
    for seg in re.split(r"[，。、；：（）()\[\]{}<>|/\s\"'“”‘’]+", desc):
        seg = seg.strip()
        if 2 <= len(seg) <= 15 and re.fullmatch(r"[\u4e00-\u9fffA-Za-z0-9_.\-]+", seg):
            out.add(seg)
    return {w for w in out if len(w) >= 2 and w not in STOP}


def overlap(a, b):
    """返回重叠对：完全相同，或一个是另一个的子串（都要求长度≥2）"""
    hits = []
    for x in sorted(a):
        for y in sorted(b):
            if x == y:
                hits.append((x, y, "完全相同"))
            elif len(x) >= 2 and len(y) >= 2 and (x in y or y in x):
                hits.append((x, y, "包含关系"))
    return hits


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=os.path.expanduser("~/.doubao/agent_mode/workspace/.user_skills"))
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    args = ap.parse_args()

    if not os.path.isdir(args.root):
        print(f"[错误] 目录不存在：{args.root}")
        return 2

    mods = []
    for d in sorted(os.listdir(args.root)):
        p = os.path.join(args.root, d, "SKILL.md")
        if os.path.isfile(p):
            name, desc = read_frontmatter(p)
            if not desc:
                print(f"[跳过] {d}：SKILL.md 没有可解析的 description")
                continue
            mods.append({"dir": d, "name": name or d, "desc": desc, "words": candidates(desc)})

    if len(mods) < 2:
        msg = f"只有 {len(mods)} 个模块，无冲突可查。"
        if args.json:
            print(json.dumps({"modules": [m["name"] for m in mods], "conflicts": [], "note": msg},
                             ensure_ascii=False, indent=2))
        else:
            for m in mods:
                print(f"[模块] {m['name']}：候选触发词 {len(m['words'])} 个")
            print(msg)
        return 0

    results = []
    conflict = False
    for i in range(len(mods)):
        for j in range(i + 1, len(mods)):
            a, b = mods[i], mods[j]
            hits = overlap(a["words"], b["words"])
            results.append({"a": a["name"], "b": b["name"], "hits": hits})
            if hits:
                conflict = True

    if args.json:
        print(json.dumps({
            "root": args.root,
            "modules": [{"name": m["name"], "candidates": len(m["words"])} for m in mods],
            "pairs": results,
            "has_conflict": conflict,
        }, ensure_ascii=False, indent=2))
    else:
        print(f"[扫描] {args.root}")
        for m in mods:
            print(f"  [模块] {m['name']}：候选触发词 {len(m['words'])} 个")
        print()
        for r in results:
            if r["hits"]:
                print(f"[冲突] {r['a']} ↔ {r['b']}：重叠 {len(r['hits'])} 处")
                for x, y, kind in r["hits"][:20]:
                    print(f"    - 「{x}」 ↔ 「{y}」（{kind}）")
                if len(r["hits"]) > 20:
                    print(f"    …… 另有 {len(r['hits']) - 20} 处")
            else:
                print(f"[通过] {r['a']} ↔ {r['b']}：无重叠")
        print()
        print("结论：" + ("发现重叠，需要改新模块的 description 再装" if conflict
                        else "无重叠，可以并存"))

    return 1 if conflict else 0


if __name__ == "__main__":
    sys.exit(main())
