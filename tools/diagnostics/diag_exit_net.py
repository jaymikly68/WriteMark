# -*- coding: utf-8 -*-
"""抓「关闭 Word 卡住的 18 秒」里它到底在等谁。

思路：开启 DNS 客户端操作日志 + 高频采样 netstat，然后点×关闭 Word，
把卡住窗口内（相对时间 >= 0 秒）的 DNS 事件和 TCP 状态全部导出来。

DNS 日志是决定性证据：如果 Word 在退出时解析某个域名并超时，
这里会留下「查询失败/超时」的事件，直接指认目标域名。
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
import re
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from watermark_tool import office_tweak as ot  # noqa: E402

DNS_LOG = "Microsoft-Windows-DnsClient/Operational"
CST = timezone(timedelta(hours=8))     # 用北京时间便于人读
HANG_LIMIT = 45.0


def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", **kw)


def iso(dt):
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def dns_set(on):
    run(["wevtutil", "cl", DNS_LOG])
    if on:
        run(["wevtutil", "sl", DNS_LOG, "/e:false"])
        r = run(["wevtutil", "sl", DNS_LOG, "/e:true", "/ms:52428800"])
        return r.returncode == 0
    run(["wevtutil", "sl", DNS_LOG, "/e:false"])
    return True


def dns_events(since_iso):
    out = run(["wevtutil", "qx", DNS_LOG, "/rd:true", "/f:text",
               "/q:*[System[TimeCreated[@SystemTime >= '%s']]]" % since_iso]).stdout
    hits, cur = [], {}
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("Event ID:"):
            if cur:
                hits.append(cur)
            cur = {"id": line.split(":", 1)[1].strip(), "raw": []}
        elif cur and line:
            cur["raw"].append(line)
    if cur:
        hits.append(cur)
    res = []
    for h in hits:
        blob = "\n".join(h["raw"])
        m = re.search(r"(?:Query Records|Name|Query Name)\s*[:\s]\s*([A-Za-z0-9.\-_]+\.[A-Za-z]{2,})",
                      blob)
        name = m.group(1) if m else ""
        status = ""
        if "Query failed" in blob or "non-existent domain" in blob.lower():
            status = "FAILED"
        res.append({"id": h["id"], "name": name, "status": status})
    return res


def tcp_snapshot():
    out = run(["netstat", "-an", "-o", "/FO", "CSV", "/NH"]).stdout
    rows = []
    for line in out.splitlines():
        p = [x.strip().strip('"') for x in line.split(",")]
        if len(p) >= 5 and p[0].upper().startswith("TCP"):
            rows.append((p[1], p[2], p[3], p[4]))
    return rows


def main():
    if not ot.app_path(ot.WORD_EXE):
        print("没找到 Word")
        return 1
    running = ot.pids_named(ot.WORD_EXE)
    if running:
        print("请先关闭所有 Word（当前 PID：%s）" % sorted(running))
        return 1

    print("开启 DNS 日志…")
    if not dns_set(True):
        print("⚠ wevtutil 未能开启 DNS 日志（可能无权限），继续尝试")

    exe = ot.app_path(ot.WORD_EXE)
    proc = subprocess.Popen([exe], close_fds=True)
    hwnd = None
    t0 = time.time()
    while time.time() - t0 < 45:
        new = ot.pids_named(ot.WORD_EXE)
        if new:
            hwnd = ot._find_window(new, ot.WORD_WINDOW_CLASS)
            if hwnd:
                break
        time.sleep(0.1)
    if not hwnd:
        proc.kill()
        print("Word 窗口没出现")
        return 1
    print("窗口已出现，静置 4 秒…")
    time.sleep(4)

    print("开始取样，并关闭 Word…")
    close_iso = iso(datetime.now(CST))
    t_close = time.time()
    u = ot._user32()
    u.PostMessageW(hwnd, 0x0010, 0, 0)

    samples = []
    hung_since = None
    while time.time() - t_close < HANG_LIMIT:
        alive = ot.pids_named(ot.WORD_EXE)
        el = time.time() - t_close
        snap = tcp_snapshot()
        samples.append((el, snap, bool(alive)))
        if hung_since is None and alive and el > 1.0:
            hung_since = el
        if not alive:
            break
        time.sleep(0.15)

    dt = time.time() - t_close
    gone = not ot.pids_named(ot.WORD_EXE)
    print("点×→进程结束：%.2fs  已退出=%s" % (dt, gone))

    print("关闭 DNS 日志…")
    dns_set(False)

    print("\n===== 卡住窗口内的 DNS 事件（关闭时刻为 0 秒）=====")
    evs = dns_events(close_iso)
    if not evs:
        print("（无）")
    for e in evs:
        print("  %-6s %-14s %s" % (e["id"], e["status"], e["name"]))

    print("\n===== 剩余 TCP 连接（卡住结束时仍挂着）=====")
    if samples:
        last = samples[-1][1]
        kept = [r for r in last if r[2] not in ("LISTENING",)]
        if not kept:
            print("（无）")
        for laddr, raddr, state, pid in sorted(set(kept))[:40]:
            print("  %-28s -> %-28s %-14s pid=%s" % (laddr, raddr, state, pid))

    print("\n===== 每 0.15s 采样里，STATE 出现过的连接数 =====")
    cnt = {}
    for _, snap, _ in samples:
        for r in snap:
            cnt[r[2]] = cnt.get(r[2], 0) + 1
    for k in sorted(cnt, key=lambda x: -cnt[x]):
        print("  %-14s %d 次" % (k, cnt[k]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
