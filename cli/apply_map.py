# -*- coding: utf-8 -*-
"""
CLI：把"哪个节点放什么内容"的映射落成专属模板包。

    python cli/apply_map.py 简历.docx --map 映射.json [--pack user_ab12cd34] [--name 显示名]

映射 JSON 的结构（与网页版模型输出同一个契约，校验规则也同一套）：

    {
      "fields":  {"name": [...], "phone": [...], "email": [...], ...},
      "slots":   {"objective": {..}, "proj_title": [...], "modules": {"nodes": [[..],[..]]},
                  "skills": {"nodes": [[..],[..]]}, "tech": [...], ...},
      "cert_text": {...},
      "edu_line":  [...]
    }

modules / skills / objective 允许"成套"（嵌套列表）也允许扁平列表——扁平的由引擎
按节点所在容器（绘图/VML/普通段落）自动分套。**不要**自己凑套数。

越界、一个下标被两个区域抢、套数不一致，都会在这一步被拦下并修，
不会静默写进 Word。
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import wizard


def _read_mapping(spec):
    if spec == "-":
        return json.loads(sys.stdin.read())
    with open(spec, encoding="utf-8") as f:
        return json.load(f)


def main():
    # 中文用 \\uXXXX 转义后再打屏：GBK 控制台上 print 中文会抛 UnicodeEncodeError
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="映射校验 + 落模板包")
    ap.add_argument("docx", help="用户的简历 .docx（会被复制成模板包里的 模板.docx）")
    ap.add_argument("--map", required=True, help="映射 JSON 文件路径，或 - 表示标准输入")
    ap.add_argument("--pack", default="", help="模板包目录名，留空自动生成 user_xxxxxxxx")
    ap.add_argument("--name", default="", help="模板显示名（一般是原文件名）")
    a = ap.parse_args()

    try:
        mapping = _read_mapping(a.map)
    except Exception as e:
        print(json.dumps({"ok": False, "error": "映射 JSON 读不进来：%s: %s"
                          % (type(e).__name__, e)}, ensure_ascii=True))
        return 2

    if not isinstance(mapping, dict):
        print(json.dumps({"ok": False,
                          "error": "映射必须是 JSON 对象，收到 "
                                   + type(mapping).__name__}, ensure_ascii=True))
        return 2

    out = wizard.save_mapping_pack(a.docx, mapping, a.name, a.pack or None)
    ok = out.get("status") == "ok"
    print(json.dumps({"ok": ok,
                      "pack": out.get("name"),
                      "message": out.get("message"),
                      "missing": out.get("missing", []),
                      "absent": out.get("absent", []),
                      "counts": out.get("counts", {}),
                      "samples": out.get("samples", {}),
                      "n_nodes": out.get("n_nodes")}, ensure_ascii=True))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
