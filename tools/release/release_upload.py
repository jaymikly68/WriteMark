#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把构建好的 exe 作为 GitHub Release 附件上传（绕开 git push，不受仓库 100MB 单文件限制）。

用法：
    python release_upload.py <exe路径> [--tag v1.5.0] [--label "WriteMark v1.5.0 安装包"]

说明：
  * Release 附件单文件上限 2GB（免费账号同样适用），比 Git 仓库 100MB 宽松得多。
  * 走 GitHub REST API，不依赖 git / credential helper —— 当本地代理把 git 的
    CONNECT 隧道掐断（502）时，这条路径仍然可用。
  * 若同名的 asset 已存在会先删除再重传，重复执行安全。
"""
import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.request

API = "https://api.github.com"
OWNER, REPO = "jaymikly68", "WriteMark"
# 注意：token 只从环境变量 GH_TOKEN 读取，绝不写进源码 —— 否则会被 GitHub Push Protection
# 的 secret scanning 判定为 secret 并直接拒绝 push。
TOKEN = os.environ.get("GH_TOKEN", "")
UA = {"User-Agent": "WriteMark-release-uploader",
      "Accept": "application/vnd.github+json"}
if TOKEN:
    UA["Authorization"] = "token " + TOKEN


def api(method, path, body=None, timeout=90):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(API + path, data=data, method=method,
                                 headers=dict(UA, **({"Content-Type": "application/json"} if body else {})))
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            # 204 No Content（如 DELETE asset）没有响应体，别去 json.load，
            # 否则会抛 JSONDecodeError —— 本项目踩过。
            body = r.read()
            return r.status, (json.loads(body) if body else None)
    except urllib.error.HTTPError as e:
        raw = e.read()[:500].decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw)
        except json.JSONDecodeError:
            return e.code, raw


def ensure_release(tag, commitish):
    st, rel = api("GET", f"/repos/{OWNER}/{REPO}/releases/tags/{tag}")
    if st == 200:
        return rel
    st, rel = api("POST", f"/repos/{OWNER}/{REPO}/releases", {
        "tag_name": tag, "target_commitish": commitish,
        "name": f"WriteMark {tag}", "draft": False, "prerelease": False,
        "body": f"详见 README 版本历史。Release 由 release_upload.py 生成（自动建 tag）。",
    })
    if st >= 400:
        raise SystemExit(f"创建 Release 失败 {st}: {rel}")
    return rel


def delete_asset(release_id, name):
    st, rel = api("GET", f"/repos/{OWNER}/{REPO}/releases/{release_id}")
    if st != 200:
        return
    for a in rel.get("assets", []):
        if a["name"] == name:
            api("DELETE", f"/repos/{OWNER}/{REPO}/releases/assets/{a['id']}")
            print(f"  已删除同名旧附件 {name}")


def upload(exe, upload_url, name, label, tries=5):
    size = os.path.getsize(exe)
    data = open(exe, "rb").read()
    print(f"  上传 {name}  {size/1048576:.1f} MB")
    for i in range(1, tries + 1):
        try:
            # upload_url 形如 ".../assets{?name,label}"，是模板不是真参数，必须裁掉
            base = upload_url.split("{?")[0]
            import urllib.parse
            q = urllib.parse.urlencode({"name": name} | ({"label": label} if label else {}))
            req = urllib.request.Request(
                base + "?" + q,
                data=data, method="PUT",
                headers=dict(UA, **{"Content-Type": "application/octet-stream",
                                    "Content-Length": str(size)}))
            with urllib.request.urlopen(req, timeout=1800) as r:
                asset = json.load(r)
            print(f"  ✓ 上传成功  id={asset['id']}  {asset.get('size',0)/1048576:.1f} MB")
            return asset
        except (urllib.error.URLError, urllib.error.HTTPError, ConnectionError) as e:
            code = getattr(e, "code", "网络错误")
            print(f"  第 {i}/{tries} 次失败（{code}），5 秒后重试…")
            time.sleep(5)
    raise SystemExit("上传失败，请检查网络后重跑（脚本可重复执行）。")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("exe", nargs="?")
    ap.add_argument("--tag", default="v1.5.0")
    ap.add_argument("--label", default="")
    ap.add_argument("--commitish",
                    default="2bb57a7c07c7012ad30028c1dba7d699772bde86")
    ap.add_argument("--body-file", help="用文件里的 Markdown 更新 Release 说明，可只传此项")
    a = ap.parse_args()

    if not TOKEN:
        raise SystemExit("缺少 GH_TOKEN。请先设置环境变量：\n"
                         '  PowerShell : $env:GH_TOKEN = "你的token"\n'
                         "  Git Bash   : export GH_TOKEN=你的token")

    rel = ensure_release(a.tag, a.commitish)

    # 只更新说明，不传附件
    if a.exe is None:
        if not a.body_file:
            raise SystemExit("请给出 exe 路径，或用 --body-file 只更新说明")
        with open(a.body_file, encoding="utf-8") as f:
            body = f.read()
        st, _ = api("PATCH", f"/repos/{OWNER}/{REPO}/releases/{rel['id']}", {"body": body})
        print(f"Release 说明已更新（{st}）：{rel['html_url']}")
        for x in rel.get("assets", []):
            print(f"  已存在附件 {x['name']}  {x['browser_download_url']}")
        return

    if not os.path.isfile(a.exe):
        raise SystemExit(f"找不到文件：{a.exe}")
    name = os.path.basename(a.exe)

    print(f"Release: {rel['html_url']}  (id={rel['id']})")
    delete_asset(rel["id"], name)
    asset = upload(a.exe, rel["upload_url"], name, a.label or name)

    print("\n附件下载链接：")
    print(f"  {asset['browser_download_url']}")
    print(f"\nRelease 页面: {rel['html_url']}")


if __name__ == "__main__":
    main()
