# -*- coding: utf-8 -*-
"""诊断 / 修复「联网时关闭 Word 卡顿十几秒」。

根因（本机实测链路）
--------------------
1. 本工具退出只要 0.09 秒，且不启动 Word；工具进程退出后再手开 Word 一样卡；
2. 从未被调用过的 PowerPoint 同样卡 18 秒 → 与本工具无关；
3. 关闭那十几秒里 WINWORD 的 CPU=0ms、磁盘 IO=0，只有几条到微软云 443 的
   ESTABLISHED 连接挂着 → 在等云端会话收尾；
4. 断网再点× → 秒退（连接立刻失败，跳过等待）；
5. Word 里存在 ConnectedOneAuthAccountId → 已登录 Microsoft 账户。

所以修复方向只有一个：**把 Office 的联网行为关掉**。
这里只写 HKCU 下的 Policies（不需要管理员权限），并完整备份，随时可还原。

本模块零第三方依赖：进程用 tasklist，窗口用 ctypes 调 user32 —— 这样打包成
exe 时不需要额外 hiddenimports，也不会漏文件。
"""
from __future__ import annotations

import ctypes
import json
import os
import sys
import subprocess
import time
import winreg
from ctypes import wintypes

HK = winreg.HKEY_CURRENT_USER
MISSING = "__MISSING__"

WORD_EXE = "WINWORD.EXE"
WORD_WINDOW_CLASS = "OpusApp"

BACKUP_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
    "WordWatermarkTool")
BACKUP_FILE = os.path.join(BACKUP_DIR, "office_policy_backup.json")
# 首次改动时把「这台机器最初的值」永久记下来。之后的 apply 都从这份快照取原始值，
# 否则反复点「一键修复」会把「上一次改完的值」当成原始值备份，导致还原失效。
SNAPSHOT_FILE = os.path.join(BACKUP_DIR, "office_original_snapshot.json")


def _snap_key_addin(it):
    return "addin|%s|%s|%s" % (it["path"], it["name"], it.get("root"))


def _snap_key_user(it):
    return "user|%s|%s" % (it["path"], it["name"])


def _load_snapshot():
    """返回快照里的 {键: 原始值} 字典（文件名只是容器，键在 "values" 里）。"""
    if not os.path.isfile(SNAPSHOT_FILE):
        return {}
    try:
        with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data.get("values", {}) or {}
    except Exception:
        return {}
    return {}


def _merge_snapshot(entries, keyfn):
    """把本次观察到的原值并入快照（只补空缺，绝不覆盖已有记录）。

    这就是「首次看到的才是原始值」的保证：第一次点修复时把 3 记下来，
    之后无论点多少次修复、还原，快照里的 3 都不会变。
    """
    snap = _load_snapshot()
    changed = False
    for it in entries:
        k = keyfn(it)
        if k in snap:
            it["old"] = snap[k]
            continue
        snap[k] = it["old"]
        changed = True
        it["old"] = snap[k]
    if changed:
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
                json.dump({"time": time.strftime("%Y-%m-%d %H:%M:%S"),
                           "values": snap}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return entries


def original_of(kind, path, name, root=None):
    """取某项设置的【原始值】：先查快照，再从备份文件兜底，都没有就返回当前值。"""
    if kind == "addin":
        k = "addin|%s|%s|%s" % (path, name, root)
    else:
        k = "user|%s|%s" % (path, name)
    snap = _load_snapshot()
    if k in snap:
        return snap[k]
    # 快照里没有，再看备份文件里是否记过（本次会话之前那一次 apply 读到的值）
    if os.path.isfile(BACKUP_FILE):
        try:
            with open(BACKUP_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            pool = list(data.get("addins", [])) + list(data.get("user", []))
            for it in pool:
                mkey = (_snap_key_addin(it) if "root" in it else _snap_key_user(it))
                if mkey == k and it.get("old") != MISSING:
                    # 快照尚未收录这个键，但备份里存的是「更早一次看到的原始值」
                    snap[k] = it["old"]
                    try:
                        with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
                            json.dump({"time": time.strftime("%Y-%m-%d %H:%M:%S"),
                                       "values": snap}, f, ensure_ascii=False, indent=2)
                    except Exception:
                        pass
                    return it["old"]
        except Exception:
            pass
    return MISSING
RESULT_FILE = os.path.join(BACKUP_DIR, "office_fix_result.txt")


def _result(ret):
    """统一返回结构：{ok, msg, need_admin}。"""
    return {"ok": ret[0], "msg": ret[1], "need_admin": ret[2]}


# --------------------------------------------------------------- 注册表基础

def office_version():
    """返回 HKCU 下存在的最高 Office 主版本分支（'16.0' / '15.0' …）。"""
    root = r"Software\Microsoft\Office"
    found = []
    try:
        k = winreg.OpenKey(HK, root)
        try:
            for i in range(winreg.QueryInfoKey(k)[0]):
                sub = winreg.EnumKey(k, i)
                if not sub.endswith(".0"):
                    continue
                try:
                    if float(sub[:-2]) >= 14:
                        found.append(sub)
                except ValueError:
                    continue
        finally:
            winreg.CloseKey(k)
    except OSError:
        pass
    found.sort(key=lambda s: float(s[:-2]))
    return found[-1] if found else "16.0"


def read_val(path, name, root=HK):
    try:
        k = winreg.OpenKey(root, path, 0, winreg.KEY_READ)
    except OSError:
        return MISSING
    try:
        v, _ = winreg.QueryValueEx(k, name)
        return v
    except OSError:
        return MISSING
    finally:
        winreg.CloseKey(k)


def write_val(path, name, value, typ=winreg.REG_DWORD):
    """写 HKCU 下的值（父键由 Windows 自动补建）。"""
    h = winreg.CreateKey(HK, path)
    try:
        winreg.SetValueEx(h, name, 0, typ, value)
    finally:
        winreg.CloseKey(h)


def write_val_root(path, name, value, root, typ=winreg.REG_DWORD):
    """写指定根键下的值（HKLM 的加载项注册在这里）。

    注意：这里刻意【不开 intermediates】再逐级 OpenKey/CloseKey —— 实测那种写法
    在这台机器上写的键不会落盘（读回来还是旧值），而 path 一次传给 CreateKey
    则由 Windows 自己补建父键，能真正写进去。
    """
    h = winreg.CreateKey(root, path)
    try:
        winreg.SetValueEx(h, name, 0, typ, value)
    finally:
        winreg.CloseKey(h)


def delete_val(path, name, root=HK):
    """删除一个值；本来就没有 / 键也不存在，都算成功（幂等）。

    还原"用户原本没设置过"的项必须走这条路 —— 如果当时是 CreataKey 补建的父键，
    单纯写一个别的值过去，反而会给 Office 留下一个错误的设置。
    """
    try:
        k = winreg.OpenKey(root, path, 0, winreg.KEY_SET_VALUE)
    except OSError:
        return True                      # 键都不存在，说明那项本来就没被我们加过
    try:
        winreg.DeleteValue(k, name)
        return True
    except FileNotFoundError:
        return True                      # 值本来就不存在
    except OSError:
        return False
    finally:
        winreg.CloseKey(k)


def write_and_verify(path, name, value, root=HK, typ=winreg.REG_DWORD, tries=6):
    """写入并回读校验，不一致就重试。

    有些环境（沙箱 / 安全软件 / 第三方服务抢着改同一条键）会让写入看起来成功
    却没落盘。实测不加校验时，禁用 Word 加载项的改动会被悄悄还原，必须重试到
    确认写入生效为止。
    """
    last = None
    for _ in range(tries):
        try:
            write_val_root(path, name, value, root, typ)
        except Exception as e:
            raise
        got = read_val(path, name, root)
        if got != MISSING and int(got) == int(value):
            return True
        last = got
        time.sleep(0.25)
    return False


def del_val(path, name):
    try:
        h = winreg.OpenKey(HK, path, 0, winreg.KEY_WRITE)
    except OSError:
        return
    try:
        winreg.DeleteValue(h, name)
    except OSError:
        pass
    finally:
        winreg.CloseKey(h)


def del_val_root(path, name, root):
    """同上，但作用于指定根键（HKLM 加载项用）。"""
    try:
        h = winreg.OpenKey(root, path, 0, winreg.KEY_WRITE)
    except OSError:
        return
    try:
        winreg.DeleteValue(h, name)
    except OSError:
        pass
    finally:
        winreg.CloseKey(h)


# --------------------------------------------------------------- 策略清单

def is_admin():
    """当前进程是否具备管理员权限（用于决定要不要走 UAC 分支）。"""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def can_write(path):
    """试着打开给定路径的写句柄，判断这台机器是否允许写该键。"""
    try:
        h = winreg.OpenKey(HK, path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
        winreg.CloseKey(h)
        return True
    except OSError:
        return False


def policy_items(ver=None, strong=False):
    """需要管理员权限的组策略项。

    键名取自微软官方文档《使用策略设置管理 Microsoft 365 企业应用版的隐私控件》：
    https://learn.microsoft.com/zh-cn/microsoft-365-apps/privacy/manage-privacy-controls

    注意：不少机器（含本机）的 HKCU\\Software\\Policies\\Microsoft\\Office 被 ACL
    保护，连管理员不提权都写不进去。所以这里只作为「可选加码」，主力放在
    user_items() —— 用户级键不需要任何权限。
    """
    ver = ver or office_version()
    base = r"Software\Policies\Microsoft\Office"
    items = [
        (base + r"\%s\Common\Privacy" % ver, "DisconnectedState", 2,
         "关闭一切需要联网的功能（连接体验）—— 最核心的一条"),
        (base + r"\%s\Common\Privacy" % ver, "usercontentdisabled", 2,
         "禁止分析文档内容的联网功能"),
        (base + r"\%s\Common\Privacy" % ver, "downloadcontentdisabled", 2,
         "禁止下载在线内容（模板 / 图片 / 字体）"),
        (base + r"\%s\Common\Privacy" % ver, "controllerconnectedservicesenabled", 2,
         "关闭其它可选连接体验"),
        (base + r"\Common\ClientTelemetry", "sendtelemetry", 3,
         "停止发送诊断遥测（3 = 必需与可选都不发）"),
    ]
    if strong:
        items.append(
            (base + r"\%s\Common\SignIn" % ver, "SignInOptions", 3,
             "禁止 Office 登录账户（用 OneDrive 云保存时需手动登录回来）"))
    return items


# --------------------------------------------------------------- 加载项

ADDIN_ROOT = r"Software\Microsoft\Office\Word\AddIns"
# 实测：这台机器上 MSOfficePLUS（微软官方 OfficePLUS 插件）随 Word 启动自动
# 加载，关闭 Word 时它在做云端收尾，把退出时间从 ~3 秒拖到 ~19 秒。
# A/B 验证：仅将其 LoadBehavior 置 0，18.70s → 2.96s。
ADDIN_TARGETS = [
    ("MSOfficePLUS", "微软官方 OfficePLUS 插件（AI 实验室 / 云素材库 / 模板）"),
    ("YunOfficeAddin.YunWordConnect", "百度网盘 Word 插件"),
    ("OneNote.WordAddinTakeNotesService", "OneNote 便笺"),
]

# LoadBehavior 取值（微软约定）
LB_DISABLE = 0    # 已断开：不加载（这就是我们要写入的值）
LB_CONNECT = 1    # 已连接，按需加载
LB_MANUAL = 2     # 已连接，但不在启动时加载
LB_AUTO = 3       # 已连接且随 Word 启动自动加载


def load_active_addins():
    """返回 [(名称, LoadBehavior, 说明, 当前是否会自动加载)]，只列注册表里存在的。"""
    out = []
    for name, desc in ADDIN_TARGETS:
        lb = read_val(ADDIN_ROOT + "\\" + name, "LoadBehavior", winreg.HKEY_CURRENT_USER)
        if lb == MISSING:
            lb = read_val(ADDIN_ROOT + "\\" + name, "LoadBehavior", winreg.HKEY_LOCAL_MACHINE)
        if lb == MISSING:
            continue
        lb = int(lb)
        out.append((name, lb, desc, lb >= 2))
    return out


def addin_items():
    """返回要写入的加载项 [(路径, 名称, 0(禁用), 说明, root)]。"""
    rows = []
    for name, desc in ADDIN_TARGETS:
        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            cur = read_val(ADDIN_ROOT + "\\" + name, "LoadBehavior", root)
            if cur == MISSING:
                continue
            path = ADDIN_ROOT + "\\" + name
            rows.append((path, name, LB_DISABLE, desc, root))
    return rows


def user_items(ver=None):
    """用户级设置项 —— 默认档。**只放没有副作用的项。**

    设计原则：卡顿的真凶已经由 COM 加载项坐实（禁用 MSOfficePLUS 后
    18.7s → 2.96s），不需要再去动账户登录状态。任何会"把人踢下线"的项
    都不该被默认打开，否则修 A 却把用户的 B 弄坏了。

    只保留 SendTelemetry：纯关遥测上传，不影响任何功能与登录。

    有副作用、需用户明确同意的项见 optional_items()。
    """
    ver = ver or office_version()
    return [
        (r"Software\Microsoft\Office\Common\ClientTelemetry", "SendTelemetry", 3,
         "停止发送诊断遥测（3 = 必需与可选都不发）—— 不影响登录和任何功能"),
    ]


def optional_items(ver=None):
    """可选加码项：**默认不写**，只在用户显式勾选时应用。

    这几条会改变 Office 的联网行为，其中 SignInOptions 会直接把已登录的账户
    踢下线（实测用户反馈"Word 账号被退出"就是这条引起的），必须让用户自己决定。
    """
    ver = ver or office_version()
    priv = r"Software\Microsoft\Office\%s\Common\Privacy" % ver
    return [
        (r"Software\Microsoft\Office\%s\Common\SignIn" % ver, "SignInOptions", 3,
         "⚠ 禁止 Office 登录账户 —— 会把已登录的账号踢下线，用 OneDrive 时每次要重新登录"),
        (priv, "DisconnectedState", 2,
         "关闭连接体验（会关掉部分需要联网的 Office 功能）"),
        (priv, "usercontentdisabled", 2,
         "禁止分析文档内容的联网功能"),
        (priv, "downloadcontentdisabled", 2,
         "禁止下载在线内容（在线模板 / 图片 / 字体）"),
        (priv, "controllerconnectedservicesenabled", 2,
         "关闭其它可选连接体验"),
    ]


def apply_optional(ver=None):
    """应用可选加码项。返回 (results, failed)。"""
    out, failed = [], []
    for path, name, want, desc in optional_items(ver):
        ok = write_and_verify(path, name, want)
        out.append((path, name, want, desc, ok))
        if not ok:
            failed.append(name)
    return out, failed


# --------------------------------------------------------------- 诊断

def diagnose():
    """只读体检。返回 dict，不改动任何东西。"""
    ver = office_version()
    ident = r"Software\Microsoft\Office\%s\Common\Identity" % ver
    acc = read_val(ident, "ConnectedOneAuthAccountId")
    logged_in = acc not in (MISSING, "", 0)

    def pack(seq, realm):
        rows, applied = [], 0
        for path, name, want, desc in seq:
            cur = read_val(path, name)
            if cur != MISSING:
                applied += 1
            rows.append({"realm": realm, "name": name,
                         "current": MISSING if cur == MISSING else cur,
                         "want": want, "desc": desc,
                         "writable": can_write(path)})
        return rows, applied

    pol, pa = pack(policy_items(ver, strong=True), "policy")
    usr, ua = pack(user_items(ver), "user")

    auto = [n for (n, lb, d, on) in load_active_addins() if on]
    return {
        "version": ver,
        "logged_in": logged_in,
        "account": (acc[:8] if isinstance(acc, str) else ""),
        "admin": is_admin(),
        "policy_writable": can_write(r"Software\Policies\Microsoft\Office\%s\Common\Privacy" % ver),
        "addins": load_active_addins(),
        "auto_addins": auto,
        "items": usr + pol,
        "applied_user": ua,
        "total_user": len(usr),
        "backup": os.path.isfile(BACKUP_FILE),
    }


# --------------------------------------------------------------- 应用 / 还原

def apply_fix(strong=False):
    """写入设置。返回 {ok, msg, need_admin}。写入前会完整备份原值。

    策略：先把【不需要管理员】的用户级项全部写下去（本机实测可写，这是主力），
    再尽力尝试组策略项；策略项被 ACL 挡掉时不整体回滚，只在报告里注明，
    因为用户级项已经生效，回滚反而会把修好的一半弄坏。
    """
    ver = office_version()
    u_before = [{"path": p, "name": n, "old": read_val(p, n)}
                for p, n, _, _ in user_items(ver)]
    a_before = _merge_snapshot(
        [{"path": p, "name": n, "root": root, "old": read_val(p, n, root)}
         for (p, n, _, _, root) in addin_items()], _snap_key_addin)
    # 用户级开关的原始值同样以「首次看到」的快照为准
    for it in u_before:
        it["old"] = original_of("user", it["path"], it["name"], None)
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        with open(BACKUP_FILE, "w", encoding="utf-8") as f:
            json.dump({"time": time.strftime("%Y-%m-%d %H:%M:%S"),
                       "ver": ver, "strong": bool(strong),
                       "user": u_before, "addins": a_before},
                      f, ensure_ascii=False, indent=2)
    except Exception as e:
        return _result((False, "备份写入失败，未改动任何设置：%s" % e, False))

    lines = ["备份：%s" % BACKUP_FILE, ""]

    # ---- 加载项：实测证明是卡顿主因，排在最前面 ----
    ok_addin, failed_addin = 0, []
    for it in a_before:
        try:
            if not write_and_verify(it["path"], it["name"], LB_DISABLE, it["root"]):
                failed_addin.append("%s（写入后校验不通过）" % it["name"])
                continue
        except Exception as e:
            failed_addin.append("%s（%s）" % (it["name"], type(e).__name__))
            continue
        ok_addin += 1
        lines.append("  [加载项] %-30s %-4s LoadBehavior %s → %d  %s"
                     % (it["name"],
                        "HKCU" if it["root"] == winreg.HKEY_CURRENT_USER else "HKLM",
                        it["old"] if it["old"] != MISSING else "?",
                        LB_DISABLE,
                        dict((n, d) for n, d in ADDIN_TARGETS).get(it["name"], "")))
    lines.append("")
    lines.append("  说明：加载项以 HKCU（用户级）的 LoadBehavior 为准，HKLM 那一份"
                 "通常只是安装时留下的，写不进去也不影响。")
    lines.append("")

    # ---- 用户级（免管理员）----
    ok_user, failed_user = 0, []
    for (path, name, val, desc), b in zip(user_items(ver), u_before):
        try:
            if not write_and_verify(path, name, val):
                failed_user.append("%s（写入后校验不通过）" % name)
                continue
        except Exception as e:
            failed_user.append("%s（%s）" % (name, type(e).__name__))
            continue
        ok_user += 1
        lines.append("  [用户级] %-36s %-14s → %-2r  %s" % (name, b["old"], val, desc))
    lines.append("")

    # ---- 组策略（可能需要管理员）----
    ok_pol, skipped_pol = 0, []
    for path, name, val, desc in policy_items(ver, strong=strong):
        try:
            write_val(path, name, val)
            ok_pol += 1
        except Exception:
            skipped_pol.append(name)

    need_admin = bool(skipped_pol) and not ok_pol
    if ok_pol:
        lines.append("  [策略级] 已写入 %d 项，进一步杜绝云端收尾" % ok_pol)
    if skipped_pol:
        lines.append("  [策略级] %d 项被系统拒绝：%s" % (len(skipped_pol), ", ".join(skipped_pol)))
        lines.append("          这台机器的 Policies 分支被 ACL 锁住，不影响已生效的部分。")
    lines.append("")
    lines.append("共改动 %d 个加载项 + %d 项开关。请完全关闭 Word"
                 "（确认任务管理器里没有 WINWORD.EXE）后重新打开再测。" % (ok_addin, ok_user))

    if not ok_addin and not ok_user and not ok_pol:
        for it in u_before:
            if it["old"] != MISSING:
                try:
                    write_val(it["path"], it["name"], it["old"])
                except Exception:
                    pass
        return _result((False, "全部写入失败：%s" % (", ".join(failed_user + failed_addin) or "未知原因"),
                        need_admin))
    if failed_addin:
        lines.append("⚠ 以下加载项未能禁用：%s" % ", ".join(failed_addin))
    if failed_user:
        lines.append("⚠ 以下开关写入失败：%s" % ", ".join(failed_user))
    return _result((True, "\n".join(lines), need_admin))


def revert_fix():
    """还原到「本工具第一次改动这台机器之前」的状态。

    原始值一律取自首次快照（SNAPSHOT_FILE），不依赖备份文件是否存在 ——
    这样即使备份被删、或者反复点过多次「一键修复」，还原出来的也始终是原样。
    原本不存在的值（原始值 = MISSING）会被直接删除。
    """
    # 从当前注册表反推「本工具动过哪些项」，再用快照里的原始值逐个写回
    targets = [(p, n, "addin", r) for (p, n, _, _, r) in addin_items()]
    targets += [(p, n, "user", None) for (p, n, _, _) in user_items()]
    targets += [(p, n, "user", None) for (p, n, _, _) in optional_items()]
    if not targets:
        return _result((False, "没有找到本工具改动过的设置（加载项或开关均不存在）", False))

    lines = ["还原到第一次改动前（依据原始值快照 %s）：" % os.path.basename(SNAPSHOT_FILE)]
    changed, failed = 0, []
    # 单项失败不该中断整批：HKLM 那一份常常没权限，而 HKCU 才是生效的那份
    for path, name, kind, root in targets:
        old = original_of(kind, path, name, root)
        try:
            if old == MISSING:
                if kind == "addin":
                    del_val_root(path, name, root)
                else:
                    del_val(path, name)
                lines.append("  删除 %-40s（原本不存在）" % name)
            else:
                if kind == "addin":
                    if not write_and_verify(path, name, old, root):
                        failed.append("%s（写入后校验不通过）" % name)
                        continue
                    lines.append("  重新启用 %-34s LoadBehavior → %s" % (name, old))
                else:
                    if not write_and_verify(path, name, old):
                        failed.append("%s（写入后校验不通过）" % name)
                        continue
                    lines.append("  还原 %-40s → %r" % (name, old))
            changed += 1
        except PermissionError as e:
            failed.append("%s（%s 无权限）" % (
                name, "HKLM" if root == winreg.HKEY_LOCAL_MACHINE else "HKCU"))
        except Exception as e:
            failed.append("%s（%s）" % (name, type(e).__name__))
    try:
        if os.path.isfile(BACKUP_FILE):
            os.remove(BACKUP_FILE)
    except Exception as e:
        lines.append("⚠ 设置已还原，但备份文件删除失败（可手动删）：%s" % e)
    lines.append("共还原 %d 项。Word 已回到改动前的状态，重启 Word 生效。" % changed)
    lines.append("提示：若还原了「关闭联网功能」类开关，Office 下次启动可能弹一次"
                 "「你的隐私由你做主」设置向导——属正常现象，手动点完即可，"
                 "注意别让它挡在 Word 后面关不掉。")
    if failed:
        lines.append("⚠ 未能处理 %d 项：%s\n  这些不影响 HKCU（用户级）那一份，"
                     "Word 实际生效的以 HKCU 为准。" % (len(failed), ", ".join(failed)))
    return _result((True, "\n".join(lines), False))


# --------------------------------------------------------------- 提权执行

def self_command():
    """构造「重新启动自己」的命令行，用于 UAC 提权。"""
    if getattr(sys, "frozen", False):
        return sys.executable, []
    main_py = os.path.abspath(sys.argv[0] or "main.py")
    return sys.executable, [main_py]


def elevate(mode):
    """用 runas 提权重开自身去执行 mode（'office-fix' / 'office-revert'）。

    返回 True 表示已成功拉起提权进程（不代表它执行成功，结果看 read_result）。
    """
    exe, pre = self_command()
    args = " ".join('"%s"' % p for p in (pre + ["--" + mode]))
    clear_result()
    # 返回值 > 32 表示成功拉起提权进程；<= 32 是失败（最常见：用户在 UAC 上点了"否"）
    # 注意：拉起成功 ≠ 执行成功，真正的结果要看 write_result 写出的文件。
    code = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", exe, args, os.getcwd(), 1)
    return code > 32, int(code)


def write_result(ok, msg):
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        with open(RESULT_FILE, "w", encoding="utf-8") as f:
            f.write(("1" if ok else "0") + "\n" + msg)
    except Exception:
        pass


def read_result():
    """读取提权子进程留下的结果。返回 (ok, msg)，没有结果文件时返回 (None, '')。"""
    if not os.path.isfile(RESULT_FILE):
        return None, ""
    try:
        with open(RESULT_FILE, "r", encoding="utf-8") as f:
            txt = f.read()
        ok, _, msg = txt.partition("\n")
        return ok.strip() == "1", msg
    except Exception:
        return None, ""


def clear_result():
    try:
        os.remove(RESULT_FILE)
    except OSError:
        pass


# --------------------------------------------------------------- 量化测试

def _user32():
    return ctypes.WinDLL("user32", use_last_error=True)


def app_path(exe_name):
    """从 App Paths 注册表取 Word 完整路径，没装则返回空串。"""
    for tmpl in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\%s",
                 r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\%s"):
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                k = winreg.OpenKey(root, tmpl % exe_name, 0, winreg.KEY_READ)
            except OSError:
                continue
            try:
                v, _ = winreg.QueryValueEx(k, "")
                if v and os.path.isfile(v):
                    return v
            except OSError:
                pass
            finally:
                winreg.CloseKey(k)
    return ""


def pids_named(exe_name):
    """用 tasklist 枚举同名进程 PID（不依赖 pywin32）。"""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq %s" % exe_name,
                              "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, timeout=15).stdout
    except Exception:
        return set()
    pids = set()
    for line in out.splitlines():
        parts = [p.strip().strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _find_window(owner_pids, win_class):
    """在给定 pid 集合里找可见的主窗口句柄。"""
    u = _user32()
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    box = {}

    def cb(hwnd, _):
        try:
            if not u.IsWindowVisible(hwnd):
                return True
            buf = ctypes.create_unicode_buffer(256)
            u.GetClassNameW(hwnd, buf, 256)
            if buf.value != win_class:
                return True
            pid = wintypes.DWORD()
            u.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value in owner_pids:
                box["h"] = hwnd
                return False
        except Exception:
            pass
        return True

    u.EnumWindows(EnumWindowsProc(cb), 0)
    return box.get("h")


def _blocking_dialogs(owner_pids):
    """找出属于这些 PID、又不是主窗口的可见弹窗（标题, 类名）。

    Word 主窗口类是 OpusApp；其余可见窗口都是挡路的模态对话框，
    典型如「你的隐私由你做主」首次设置向导、"无法关闭 Microsoft Word，
    因为有一个对话框处于打开状态" 的错误弹窗——它们会让 WM_CLOSE 静默失效，
    程序化关闭会一直等到超时。
    """
    u = _user32()
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    found = []

    def cb(hwnd, _):
        try:
            if not u.IsWindowVisible(hwnd):
                return True
            pid = wintypes.DWORD()
            u.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value not in owner_pids:
                return True
            cls = ctypes.create_unicode_buffer(256)
            u.GetClassNameW(hwnd, cls, 256)
            if cls.value == WORD_WINDOW_CLASS:
                return True  # 主窗口，不算挡路
            # 需要 400×150 以上的窗口，过滤掉 tooltip 之类的小物件
            rc = wintypes.RECT()
            if u.GetWindowRect(hwnd, ctypes.byref(rc)):
                if (rc.right - rc.left) < 200 or (rc.bottom - rc.top) < 100:
                    return True
            ttl = ctypes.create_unicode_buffer(256)
            u.GetWindowTextW(hwnd, ttl, 256)
            found.append((ttl.value.strip() or "(无标题)", cls.value))
        except Exception:
            pass
        return True

    u.EnumWindows(EnumWindowsProc(cb), 0)
    return found


def benchmark(settle=4.0, open_timeout=45, close_timeout=90):
    """量化「点× → WINWORD.EXE 进程结束」的耗时（秒）。

    返回 (秒数或 None, 说明文字)。None 表示这台机器没法测（没装 Word / 已有实例）。
    """
    exe = app_path(WORD_EXE)
    if not exe:
        return None, "这台机器上没找到 Word"
    before = pids_named(WORD_EXE)
    if before:
        return None, "已存在运行中的 Word（PID %s），请先全部关闭再测" % sorted(before)

    u = _user32()
    proc = subprocess.Popen([exe], close_fds=True)
    try:
        hwnd = None
        t0 = time.time()
        while time.time() - t0 < open_timeout:
            new = pids_named(WORD_EXE) - before
            if new:
                hwnd = _find_window(new, WORD_WINDOW_CLASS)
                if hwnd:
                    break
            time.sleep(0.1)
        if not hwnd:
            proc.kill()
            return None, "Word 窗口没有出现"
        opened = time.time() - t0

        time.sleep(settle)
        new_pids = pids_named(WORD_EXE) - before

        t1 = time.time()
        u.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
        dt = None
        block_note = ""
        last_check = 0.0
        while time.time() - t1 < close_timeout:
            if not (pids_named(WORD_EXE) - before) & new_pids:
                dt = time.time() - t1
                break
            # 每 0.5s 查一次有没有弹窗挡路（模态对话框会让 WM_CLOSE 静默失效）
            now = time.time()
            if now - last_check >= 0.5:
                last_check = now
                blks = _blocking_dialogs(new_pids)
                if blks:
                    block_note = "；发现挡路弹窗：%s" % "、".join(
                        repr(t) for t, _ in blks[:3])
                    break
            time.sleep(0.05)
    finally:
        try:
            proc.wait(3)
        except Exception:
            proc.kill()

    if dt is None:
        if block_note:
            return float("inf"), "Word 被弹窗挡住、无法程序化关闭%s。" % block_note + \
                   " 请手动关掉那个弹窗（如首次启动的隐私设置向导）再测"
        return float("inf"), "超过 %.0f 秒仍未退出%s" % (close_timeout, block_note)
    verdict = "秒退" if dt < 2 else ("仍慢" if dt > 5 else "尚可")
    return dt, "窗口 %.1fs 出现，点×→进程结束 %.2fs（%s）" % (opened, dt, verdict)
