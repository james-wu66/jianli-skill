# -*- coding: utf-8 -*-
"""
CLI：Agent（或模型）写好的定制结果 JSON → 走引擎保底 → 在模板上原地换字出 Word。

    python cli/postprocess.py --result 结果.json
    python cli/postprocess.py -  --profile 张三.json --template user_ab12cd34 --out 我的简历.docx

这是网页版 `app.py` 里"拿到模型输出之后"那条路的命令行版，**复用同一批函数**：

    generate._extract_json  → 容错解析（含 markdown 包裹、截断恢复）+ _fill 补默认
                              + _pad_tech 按 fit 方向从档案真实技术栈里保底补齐
    template_export.export_from_result → 在模板 .docx 上只换文字节点，版式不动

所以 skill 里不需要、也不应该重写这些逻辑。Agent 只负责"产出那份 JSON"，
剩下的确定性活儿全交给这里。

结果 JSON 的字段契约见 generate.py 的 _json_schema()（条目数由 slot_counts() 按模板决定）：
objective / fit / project_title /
project_role / project_points / tech_stack / skill_lines / project_links / match /
rewrite_notes。缺字段不报错，会用档案里的基础素材补，但**绝不新增档案里没有的事实**。
"""

import argparse
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generate
import paths
import profile_store
import template_export


def _read_text(spec):
    if spec == "-":
        return sys.stdin.read()
    with open(spec, encoding="utf-8") as f:
        return f.read()


def main():
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="结果 JSON → 保底补齐 → 出 Word")
    ap.add_argument("--result", "-r", default="-", help="结果 JSON 文件，或 - 读标准输入")
    ap.add_argument("--profile", default="", help="档案文件名（默认当前激活档案）")
    ap.add_argument("--template", default="", help="模板包目录名（默认取档案里的 template）")
    ap.add_argument("--out", default="", help="输出 docx 路径（默认 output/ 下带时间戳）")
    a = ap.parse_args()

    pname = a.profile
    if pname and not pname.endswith(".json"):
        pname += ".json"
    profile = profile_store.load_profile(pname or None)
    if not profile:
        print(json.dumps({"ok": False,
                          "error": "没有个人档案，profiles/ 是空的或还没设激活档案"},
                         ensure_ascii=True))
        return 3

    try:
        raw = _read_text(a.result)
    except Exception as e:
        print(json.dumps({"ok": False, "error": "结果读不进来：%s: %s"
                          % (type(e).__name__, e)}, ensure_ascii=True))
        return 2

    # 与网页版同一条解析+兜底路径
    obj = generate._extract_json(raw, profile)
    if not isinstance(obj, dict):
        print(json.dumps({"ok": False,
                          "error": "结果不是合法 JSON，也没能从截断内容里恢复出来。"
                                   "请重发一次，或把缺的字段补上"},
                         ensure_ascii=True))
        return 2
    obj.setdefault("project_points", [])
    obj.setdefault("skill_lines", [])
    obj.setdefault("tech_stack", [])
    obj.setdefault("match", {})
    obj.setdefault("rewrite_notes", "")

    fit = (obj.get("fit") or "全栈").strip()
    who = ((profile.get("basic") or {}).get("name") or "简历").strip()
    out_path = a.out or os.path.join(
        paths.output_dir(),
        "%s简历_%s_%s.docx" % (who, fit, time.strftime("%m%d_%H%M%S")))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)

    template_dir = a.template or profile.get("template")
    match = obj.get("match") or {}
    try:
        template_export.export_from_result(obj, out_path, profile=profile,
                                           template_dir=template_dir)
    except Exception as e:
        print(json.dumps({"ok": False,
                          "error": "Word 导出失败：%s: %s" % (type(e).__name__, e),
                          "template": template_dir}, ensure_ascii=True))
        return 1

    print(json.dumps({
        "ok": True,
        "out": out_path.replace("\\", "/"),
        "template": template_dir or template_export.DEFAULT_TEMPLATE,
        "fit": fit,
        "slots": obj.get("_slots") or {},
        "slot_gap": obj.get("_slot_gap") or {},
        "project_points": len(obj["project_points"]),
        "skill_lines": len(obj["skill_lines"]),
        "tech_stack": obj["tech_stack"],
        "score": match.get("score"),
        "recommend": match.get("recommend"),
        "requirements": match.get("requirements") or [],
        "summary": match.get("summary"),
        "rewrite_notes": obj.get("rewrite_notes"),
    }, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
