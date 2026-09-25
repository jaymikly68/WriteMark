# -*- coding: utf-8 -*-
"""把本地构建好的 exe 发布为 GitHub Release 附件。

只上传「实测可用」的版本：报错的（打不开 / 会踢 Word 账号）一律跳过。
用法：python _release.py
token 从环境变量 GITHUB_TOKEN 读，不写进文件。
"""
import json
import os
import sys
import time
import urllib.error
import urllib.request

API = "https://api.github.com"
REPO = "jaymikly68/WriteMark"

# (tag, 本地文件, 上传文件名, release 说明)
PLAN = [
    ("v1.1.5", "dist/WordWatermark.exe", "WordWatermark-v1.1.5.exe",
     "**本分支最终可用版本。**\n\n"
     "**修复两个我自己引入的坑**\n"
     "1. `gui.py` 漏写 `import sys`——`main()` 一进就 `NameError`，双击即闪退"
     "（就是上一版「程序打不开」的原因）\n"
     "2. 「一键修复」默认写入 `SignInOptions=4`（禁止 Office 登录任何账户），"
     "会把已登录的 Word 账号踢下线。真凶已定位为 COM 加载项，默认档不再碰登录状态，"
     "改为只禁加载项 + 关遥测；会踢账号、会关联网功能的项目拆成独立的「可选加码」按钮，"
     "需显式二次确认\n\n"
     "**其他改进**\n"
     "- 启动器新增 payload 指纹校验：本地运行体目录原来按版本号复用，"
     "版本号不变就永不重新解压，导致「发了新 exe 用户却还在跑旧代码」\n"
     "- 「Word 秒退」测速会检测挡路弹窗，不再对模态对话框干等 90 秒\n"
     "- 清掉上一版遗留的联网开关，让 Office 回到默认行为"),
    ("v1.1.3", "dist/WordWatermark_prev_0924_031555.exe", "WordWatermark-v1.1.3.exe",
     "**彻底不再启动 Word**：字体列表改为直接读 Windows 字体注册表 + 中文本地化映射表，"
     "读字体从 5.5s（且会临时起一个 Word）降到 0.006s、零副作用，"
     "中文字体名（宋体/微软雅黑/仿宋/等线…27 个）依然可用。\n\n"
     "新增 `parse_color` 兜底（传 `#FF0000` 这类字符串不再崩溃）。\n\n"
     "注：本版已实测可正常启动与使用；它已包含「启动器 + 本地常驻运行体」打包方式，"
     "关闭后 0.02 秒退出。"),
]


TOKEN = os.environ.get("GITHUB_TOKEN")
if not TOKEN:
    sys.exit("缺少环境变量 GITHUB_TOKEN")


def req(url, data=None, headers=None, method=None, timeout=900):
    """用标准库发请求，返回 (status, json文本或None)。"""
    h = {"Authorization": "Bearer %s" % TOKEN,
         "Accept": "application/vnd.github+json",
         "User-Agent": "WriteMark-release-script"}
    h.update(headers or {})
    body = None if data is None else (
        json.dumps(data).encode("utf-8") if not isinstance(data, bytes) else data)
    r = urllib.request.Request(url, data=body, headers=h, method=method)
    try:
        with urllib.request.urlopen(r, timeout=timeout) as resp:
            raw = resp.read()
            try:
                return resp.status, json.loads(raw.decode("utf-8"))
            except Exception:
                return resp.status, raw.decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, raw


def create_release(tag, name, body):
    code, res = req(API + "/repos/%s/releases" % REPO, {
        "tag_name": tag, "name": name, "body": body,
        "draft": False, "prerelease": False,
    }, {"Content-Type": "application/json"}, method="POST")
    if code == 201:
        return res["id"]
    if code == 422:
        errs = res.get("errors", []) if isinstance(res, dict) else []
        print("  tag 已存在，复用：%s" % (errs[0]["message"] if errs else res)[:140])
        code, res = req(API + "/repos/%s/releases/tags/%s" % (REPO, tag), timeout=60)
        return res["id"]
    sys.exit("创建 release %s 失败：%s %s" % (tag, code, res if not isinstance(res, dict) else res))


def upload(rid, path, asset_name):
    url = ("https://uploads.github.com/repos/%s/releases/%d/assets?name=%s"
           % (REPO, rid, asset_name))
    with open(path, "rb") as fh:
        code, res = req(url, fh.read(), {
            "Content-Type": "application/octet-stream",
            "Content-Length": str(os.path.getsize(path)),
        }, method="POST")
    if code in (201, 200):
        print("  上传成功：%s（%.1f MB）" % (
            asset_name, os.path.getsize(path) / 1048576.0))
        print("     %s" % res["browser_download_url"])
        return True
    print("  上传失败 %s：%s" % (code, res if not isinstance(res, dict)
                                else res.get("message", res))[:300])
    return False


def main():
    # PLAN 里的路径是相对**项目根**的（本脚本位于 tools/build/）
    base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    ok = True
    for tag, rel, asset, body in PLAN:
        path = os.path.join(base, rel)
        if not os.path.isfile(path):
            print("[跳过] %s —— 本地找不到 %s" % (tag, rel))
            ok = False
            continue
        print("[%s] %s -> %s" % (tag, os.path.basename(path), asset))
        rid = create_release(tag, tag, body)
        if not upload(rid, path, asset):
            ok = False
        time.sleep(1)
    print("\n全部完成" if ok else "\n有项目未完成")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
