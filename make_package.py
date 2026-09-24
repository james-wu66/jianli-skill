# -*- coding: utf-8 -*-
"""
一键打包分发：python make_package.py
产出 dist/简历定制器_<日期>.zip，可以直接发给别人。

原则：
  白名单打包——只打进"运行必需"的文件，个人数据/文档/测试一概不带。
  key 硬校验——把本机 llm_config.py 和 profiles/api_key.json 里的真实 key 读出来，
  逐个确认它们【没有】出现在任何打包文件里，发现即报错拒绝出包。
"""

import os
import re
import shutil
import sys
import zipfile

# cmd 默认代码页是 936(GBK)，✅/❌ 这类字符编不出来会抛 UnicodeEncodeError：
# 打包其实已经成功，却回你一屏 traceback 让人以为失败了；更要命的是
# "检测到 key 泄漏、拒绝出包"那两条提示也会一起变成看不懂的乱码。
# errors="replace" 让这类字符降级成 ? —— UTF-8 终端（Git Bash / VS Code）照常显示。
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass

APP = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(APP, "dist")
PKG_NAME = "简历定制器"
PKG_DIR = os.path.join(DIST, PKG_NAME)

# ---- 白名单：打进包里的文件 ----
FILES = [
    "app.py", "generate.py", "llm_client.py", "profile_store.py",
    "resume_parser.py", "template_export.py", "wizard.py", "launch.py",
    "index.html", "启动.bat", "README.txt",
]
DIRS = [
    ("assets/templates/classic_blue", ["模板.docx", "template.json"]),
]

# llm_config.py 故意不打包：收包人通过网页「设置」页填自己的 key
# （llm_client 对它的缺失有 try/except 兜底）


def _collect_keys():
    """把本机所有可能的真实 key 收集起来，供泄漏检查。"""
    keys = set()
    for path in (os.path.join(APP, "llm_config.py"),
                 os.path.join(APP, "profiles", "api_key.json")):
        try:
            text = open(path, encoding="utf-8").read()
            keys.update(re.findall(r"sk-[A-Za-z0-9_\-]{16,}", text))
        except OSError:
            pass
    return keys


def _scan_leak(pkg_dir, secrets):
    """递归扫描包内所有文本文件，发现 key 泄漏立即抛错。"""
    for root, _dirs, files in os.walk(pkg_dir):
        for fn in files:
            p = os.path.join(root, fn)
            if fn.lower().endswith((".docx", ".png", ".jpg", ".ico")):
                continue
            try:
                text = open(p, encoding="utf-8", errors="ignore").read()
            except OSError:
                continue
            for sk in secrets:
                if sk in text:
                    raise SystemExit(f"❌ 拒绝出包：检测到真实 key 泄漏到 {p}")
            if re.search(r"sk-[A-Za-z0-9_\-]{16,}", text):
                raise SystemExit(f"❌ 拒绝出包：{p} 里发现疑似 key 的字符串")


def main():
    secrets = _collect_keys()
    print(f"key 泄漏检查样本：{len(secrets)} 个（来自 llm_config / api_key.json）")

    # 干净重建
    if os.path.exists(PKG_DIR):
        shutil.rmtree(PKG_DIR)
    os.makedirs(PKG_DIR)

    for f in FILES:
        src = os.path.join(APP, f)
        if not os.path.isfile(src):
            raise SystemExit(f"❌ 白名单文件缺失：{f}（先做完 README 再打包）")
        shutil.copyfile(src, os.path.join(PKG_DIR, f))
    for src_dir, keep in DIRS:
        dst = os.path.join(PKG_DIR, src_dir)
        os.makedirs(dst, exist_ok=True)
        for fn in keep:
            shutil.copyfile(os.path.join(APP, src_dir, fn), os.path.join(dst, fn))

    # profiles 目录只留空壳（收包人自己建档案/key）
    os.makedirs(os.path.join(PKG_DIR, "profiles"), exist_ok=True)
    with open(os.path.join(PKG_DIR, "profiles", "active.txt"), "w", encoding="utf-8") as f:
        f.write("")

    _scan_leak(PKG_DIR, secrets)

    zip_base = os.path.join(DIST, f"{PKG_NAME}_v1.0")
    if os.path.exists(zip_base + ".zip"):
        os.remove(zip_base + ".zip")
    with zipfile.ZipFile(zip_base + ".zip", "w", zipfile.ZIP_DEFLATED) as z:
        for root, _dirs, files in os.walk(PKG_DIR):
            for fn in files:
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, DIST)
                z.write(full, os.path.join(PKG_NAME, os.path.relpath(full, PKG_DIR)))

    n = sum(len(f) for _r, _d, f in os.walk(PKG_DIR))
    print(f"✅ 打包完成：{zip_base}.zip（{n} 个文件）")
    print("包内清单：")
    for root, _dirs, files in os.walk(PKG_DIR):
        for fn in sorted(files):
            rel = os.path.relpath(os.path.join(root, fn), PKG_DIR)
            print("   ", rel)


if __name__ == "__main__":
    main()
