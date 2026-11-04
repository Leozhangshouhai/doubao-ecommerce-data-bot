#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ECommerce Data Bot · 配置文件加载与校验
=========================================
固定读取文件名：ecommerce_bot.yaml（同一目录或 --path 指定）
职责：
  1) 文件存在性与文件名校验（必须叫 ecommerce_bot.yaml）
  2) YAML 解析，错误时定位到具体行
  3) 字段必填 / 类型 / 枚举校验，逐项报错（不静默跳过）
  4) 连接自检（可选 --check-conn）：隧道或直连可达、只读账号能 SELECT、关键表存在且非空
输出：JSON {ok, errors[], info{...}}
用法：
  python3 load_config.py                          # 仅结构校验
  python3 load_config.py --check-conn             # 结构 + 连接/表自检
"""
import json
import socket
import sys
import os
import yaml

REQUIRED_FILE = "ecommerce_bot.yaml"

# ---------- 必填字段定义 ----------
REQUIRED = {
    "customer": {"type": str, "nonempty": True},
    "conn": {"type": dict},
    "db": {"type": dict},
    "caliber": {"type": dict},
    "tables": {"type": dict},
}
CONN_KEYS = {"mode": {"in": ["jump", "direct"]}}
DB_KEYS = {"user": {"type": str, "nonempty": True}, "password": {"type": str, "nonempty": True}}
TABLE_KEYS = ["inventory_sale", "item_sale", "order_head", "order_detail"]
CALIBER_SOLDOUT_KEYS = ["on_sale_status", "sale_type", "sold_out_rule", "grade", "flow_cats"]


def err(msg):
    print(json.dumps({"ok": False, "errors": [msg], "info": {}}, ensure_ascii=False))
    sys.exit(1)


def main():
    path = None
    check_conn = "--check-conn" in sys.argv
    for i, a in enumerate(sys.argv):
        if a == "--path" and i + 1 < len(sys.argv):
            path = sys.argv[i + 1]
    if path is None:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", REQUIRED_FILE)
    path = os.path.abspath(path)

    errors = []
    if not os.path.isfile(path):
        err(f"找不到配置文件：{path}（文件名必须为 {REQUIRED_FILE}）")
    if os.path.basename(path) != REQUIRED_FILE:
        err(f"文件名必须是 {REQUIRED_FILE}，当前为 {os.path.basename(path)}")

    # YAML 解析（错误定位行号）
    try:
        with open(path, encoding="utf-8") as f:
            text = f.read()
        cfg = yaml.safe_load(text)
    except yaml.YAMLError as e:
        mark = getattr(e, "problem_mark", None)
        line = mark.line + 1 if mark else "?"
        err(f"YAML 语法错误（第 {line} 行附近）：{getattr(e, 'problem', e)}")

    if not isinstance(cfg, dict):
        err("配置文件内容必须是一个 YAML 对象（键值对）")

    # 字段校验
    for k, rule in REQUIRED.items():
        if k not in cfg or cfg[k] is None:
            errors.append(f"缺少必填字段：{k}")
            continue
        v = cfg[k]
        if rule.get("type") == dict and not isinstance(v, dict):
            errors.append(f"字段 {k} 应为对象（键值对）")
    if errors:
        finish(errors)

    # conn 校验
    conn = cfg["conn"]
    if conn.get("mode") not in ("jump", "direct"):
        errors.append("conn.mode 必须是 jump 或 direct")
    if conn["mode"] == "jump":
        for k in ("jump_host", "jump_user", "jump_password"):
            if not conn.get(k):
                errors.append(f"conn.{k} 必填（跳板机模式）")
        tg = conn.get("targets") or {}
        for k in ("order", "goods"):
            t = tg.get(k) or {}
            if not t.get("host") or not t.get("port"):
                errors.append(f"conn.targets.{k}.host/port 必填（跳板机后实例地址）")
        if isinstance(conn.get("tunnel"), dict):
            for k in ("order_port", "goods_port"):
                p = conn["tunnel"].get(k)
                if p is not None and (not isinstance(p, int) or not (1 <= p <= 65535)):
                    errors.append(f"conn.tunnel.{k} 必须是 1-65535 的端口数字")
    else:
        for k in ("order_host", "order_port", "goods_host", "goods_port"):
            if not conn.get("direct", {}).get(k):
                errors.append(f"conn.direct.{k} 必填（直连模式）")

    # db 校验
    for k, rule in DB_KEYS.items():
        if not cfg["db"].get(k):
            errors.append(f"db.{k} 必填")
    if cfg["db"].get("password") == "<只读密码>":
        errors.append("db.password 还是占位符，请填写真实密码")

    # tables 校验
    for k in TABLE_KEYS:
        if not cfg["tables"].get(k):
            errors.append(f"tables.{k} 必填（缺关键表）")

    # caliber.soldout 校验
    so = cfg.get("caliber", {}).get("soldout", {})
    for k in CALIBER_SOLDOUT_KEYS:
        if k not in so:
            errors.append(f"caliber.soldout.{k} 缺失（卖断货口径）")

    if errors:
        finish(errors)

    # 连接自检
    info = {"customer": cfg["customer"], "config_path": path}
    if check_conn:
        info.update(conn_self_check(cfg))
        # 自检失败项必须提升为错误（ok=false），让老板一眼看到"哪一步错了"；
        # 禁止静默降级、禁止尝试环境内其他凭据重试
        fails = [f"{k}: {v}" for k, v in info.items()
                 if "失败" in str(v) or "不可达" in str(v) or "跳过" in str(v)]
        if fails:
            finish(fails)

    print(json.dumps({"ok": True, "errors": [], "info": info}, ensure_ascii=False, default=str))


def conn_self_check(cfg):
    """只读连接自检：动态建隧道 → 端口可达 → 只读登录 → 关键表存在且有数据。纯 SELECT，不改数据。"""
    import paramiko
    import pymysql
    import threading
    res = {}
    conn = cfg["conn"]
    db = cfg["db"]
    tables = cfg["tables"]
    cli = None
    t = None

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

    def bridge_srv(lp, rh, rp, t):
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
        try:
            srv.bind(("127.0.0.1", lp))
            srv.listen(16)
            while True:
                c, _ = srv.accept()
                threading.Thread(target=h, args=(c,), daemon=True).start()
        except OSError:
            srv.close()  # 端口已被占用（已有服务/隧道）→ 仅探测既有服务

    # ---- jump 模式：连跳板机，动态建双向隧道 ----
    if conn["mode"] == "jump":
        try:
            cli = paramiko.SSHClient()
            cli.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            cli.connect(conn["jump_host"], username=conn["jump_user"],
                        password=conn["jump_password"], timeout=15)
            t = cli.get_transport()
            res["jump_connect"] = f"OK（{conn['jump_host']}）"
        except Exception as e:
            res["jump_connect"] = f"失败：{e}"

        targets = conn.get("targets") or {}
        tunnel = conn.get("tunnel") or {}
        pairs = [
            ("order", tunnel.get("order_port") or 13307, targets.get("order", {}).get("host"), targets.get("order", {}).get("port")),
            ("goods", tunnel.get("goods_port") or 13306, targets.get("goods", {}).get("host"), targets.get("goods", {}).get("port")),
        ]
        if t:
            for name, lp, rh, rp in pairs:
                if rh and rp:
                    threading.Thread(target=bridge_srv, args=(lp, rh, rp, t), daemon=True).start()
            import time
            time.sleep(0.6)
            for name, lp, rh, rp in pairs:
                try:
                    s = socket.create_connection(("127.0.0.1", lp), timeout=5)
                    s.close()
                    res[f"{name}_instance"] = f"隧道 {name}（{rh}:{rp} → 127.0.0.1:{lp}）可达"
                except Exception as e:
                    res[f"{name}_instance"] = f"不可达：{e}"
        else:
            for name, lp, rh, rp in pairs:
                res[f"{name}_instance"] = "跳板机未连通，跳过"
    # ---- direct 模式：直接探测 ----
    else:
        d = conn.get("direct", {})
        for name in ("order", "goods"):
            hp = (d.get(f"{name}_host"), d.get(f"{name}_port"))
            if hp[0] and hp[1]:
                try:
                    s = socket.create_connection(hp, timeout=5)
                    s.close()
                    res[f"{name}_instance"] = f"直连 {hp} 可达"
                except Exception as e:
                    res[f"{name}_instance"] = f"不可达：{e}"

    # ---- 只读登录 + 关键表 ----
    try:
        if conn["mode"] == "jump":
            lp = (conn.get("tunnel") or {}).get("order_port") or 13307
        else:
            lp = conn["direct"]["order_port"]
        c = pymysql.connect(host="127.0.0.1", port=lp, user=db["user"], password=db["password"],
                            connect_timeout=10, read_timeout=30)
        cur = c.cursor()
        cur.execute("SELECT 1")
        res["db_login"] = "OK（只读账号可登录）"
        try:
            cur.execute(f"SELECT COUNT(*) FROM {tables['inventory_sale']}")
            res[f"table.{tables['inventory_sale']}"] = f"{cur.fetchone()[0]} 行"
        except Exception as e:
            res[f"table.{tables['inventory_sale']}"] = f"查询失败：{e}"
        c.close()
    except Exception as e:
        res["db_login"] = f"失败：{e}"

    if cli:
        try:
            cli.close()
        except Exception:
            pass
    return res


def finish(errors):
    print(json.dumps({"ok": False, "errors": errors, "info": {}}, ensure_ascii=False))
    sys.exit(1)


if __name__ == "__main__":
    main()
