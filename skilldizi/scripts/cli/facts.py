# -*- coding: utf-8 -*-
"""
CLI：把当前档案渲染成引擎一直在用的那三份提示词文本。

    python cli/facts.py [--profile 张三.json]

输出三段，全部来自 generate.py 里现成的函数，不在这里重新拼：

    1) 事实清单   generate._profile_to_text   —— 唯一信息源，姓名/电话/院校/课程/证书都在里面
    2) 个人红线   generate._rules_text        —— profile.rules（表述红线 + 标签黑白名单）
    3) 评分短板   generate._scoring_text      —— profile.scoring.facts，防止打分虚高

为什么要这个 CLI：这三段以前是引擎拼给模型的。skill 形态下"模型"换成了宿主 Agent，
但如果让 Agent 自己去啃 profiles/*.json 的原始结构，键名、层级、哪些字段该进清单
全都得它自己猜——猜错就是把别人的事实写进你的简历。渲染逻辑只留一份。
"""

import argparse
import io
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generate
import profile_store


def main():
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="档案 → 提示词文本")
    ap.add_argument("--profile", default="", help="档案文件名，留空用当前激活档案")
    ap.add_argument("--template", default="",
                    help="这次要写入的模板包名；留空用档案里的 template 字段")
    ap.add_argument("--out", default="", help="写进这个文件（UTF-8）；不填就打屏")
    a = ap.parse_args()

    pname = a.profile
    if pname and not pname.endswith(".json"):
        pname += ".json"
    profile = profile_store.load_profile(pname or None)
    if not profile:
        print(json.dumps({"ok": False,
                          "error": "没有个人档案：profiles/ 是空的或还没设激活档案。"
                                   "先建档案再定制。"}, ensure_ascii=True))
        return 3

    tpl = a.template or profile.get("template") or "classic_blue"
    slots = generate.slot_counts(profile, tpl)
    buf = io.StringIO()
    buf.write("【模板槽位】项目要点 %d 行 ｜ 技能 %d 行 ｜ 模板包 %s\n"
              "（产出条数必须等于这个数：多了会被截掉，少了会留空行，不要按 6/5 想当然）\n\n"
              % (slots["modules"], slots["skills"], tpl))
    buf.write("=== 用户唯一的真实简历事实清单（唯一信息源，禁止超出范围编造）===\n\n")
    buf.write(generate._profile_to_text(profile))
    rules = generate._rules_text(profile)
    if rules.strip():
        buf.write("\n=== 以下来自档案 rules，个人红线，同样必须遵守 ===\n\n")
        buf.write(rules.rstrip() + "\n")
    scoring = generate._scoring_text(profile)
    if scoring.strip():
        buf.write("\n=== 评分短板（打匹配分时必须如实扣，禁止虚高）===\n\n")
        buf.write(scoring.strip() + "\n")
    text = buf.getvalue()

    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(text)
        print(json.dumps({"ok": True, "out": a.out.replace("\\", "/"),
                          "chars": len(text),
                          "profile": profile.get("_file"),
                          "slots": slots,
                          "has_rules": bool(rules.strip()),
                          "has_scoring": bool(scoring.strip())}, ensure_ascii=True))
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
