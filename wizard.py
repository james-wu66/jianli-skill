# -*- coding: utf-8 -*-
"""
首次配置向导（M2，2026-09-15 新增）：
  1. extract_profile：用大模型把用户上传的简历原文整理成结构化档案（profile 草稿）
     ——只提取简历里真实存在的内容，禁止编造。
  2. map_template：把用户上传的 docx 当作他的专属导出模板——dump 文字节点后
     用大模型生成"哪个节点换什么内容"的映射配置（= 自动生成 template.json），
     生成时版式和用户自己的简历一模一样。
  3. 映射经过严格校验（下标越界/一址多写/两套槽数不一致都会被修掉），
     修完不可用就回退内置模板 classic_blue。

两个提示词的设计要点见各自函数内的 docstring。
"""

import json
import os
import re
import shutil
import uuid

from llm_client import call_model
from resume_parser import _extract_lines
import template_export

APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = template_export.TEMPLATES_DIR

# 提取档案：输出约 1~3 千 token，12000 足够且不触发长思考
_EXTRACT_MAX_TOKENS = 12000
# 模板映射：输出 1~2 千 token
_MAP_MAX_TOKENS = 8000

# ---------------------------------------------------------------------------
# 提示词一：简历原文 → 个人档案草稿
# 设计要点：结构 = profile JSON 的骨架；逐字段说明取法；把"禁止编造"放在最前面；
# 拿不准就留空，后面有人工确认表单兜底。
# ---------------------------------------------------------------------------
_EXTRACT_SYSTEM = """你负责把用户的简历原文整理成结构化档案（JSON）。
铁律：只允许提取简历里真实出现的内容，禁止编造、禁止推测、禁止补充简历里没有的任何信息。
拿不准或简历里没有的字段，一律用空字符串或空列表，绝不要填占位内容。

输出 JSON 结构（字段名严格一致）：
{
 "basic": {"name":"","ethnic":"","phone":"","email":"","birth":"","hukou":"","degree":"","github":"","demo":""},
 "education": {"period":"","school":"","major":"","courses":""},
 "objective_default": "",
 "objective_candidates": [],
 "projects": [{
   "title":"","role":"","period":"",
   "links":{"github":"","demo":""},
   "intro":"",
   "modules":[{"name":"","core":false,"default":""}],
   "tech_stack":[]
 }],
 "skills": {"分组名": ["条目"]},
 "certificates": []
}

取法说明：
- basic.name=姓名；ethnic=民族；phone=电话；email=邮箱；birth=出生年月；hukou=户籍/籍贯；
  degree=本科/硕士/大专这类；github=个人 GitHub 主页；demo=项目演示地址。
- education：起止时间、学校、专业、主修课程。
- objective_default=求职意向原文；objective_candidates=简历里出现过的可选意向。
- projects：简历里的每个项目一条；modules 取项目经历里的每条要点
  （形如「●模块名：内容」，name=模块名，default=简历原句，core=是否重点），最多 10 条；
  tech_stack=该项目实际用到的技术名词。
- skills：按简历的技能分组原样提取；没有分组就全放进「技能」一组。
- certificates：所有证书、奖项；没有就空列表。

只输出 JSON，不要 markdown 包裹，不要任何多余文字。"""

# ---------------------------------------------------------------------------
# 提示词二：文字节点清单 → 模板映射配置
# 设计要点：先讲引擎原理（原地换文字、版式不动），再给目标 JSON 骨架，
# 然后是 5 条判断规则（双份内容/标签值分离/宁缺勿错/禁止一址多写/只输出 JSON）。
# ⚠️ 节点清单必须放 user 消息——商汤 API 对"只有 system 没有 user"的请求直接 400
#    （2026-09-15 实测踩过：报 inference request is invalid）。
# ---------------------------------------------------------------------------
_MAP_SYSTEM = """你在为「简历模板替换引擎」分析一份 docx 简历的内部文字节点。
引擎原理：生成新简历时，把这些节点的文字【原地替换】成新内容，版式、形状、图片完全不动。
你的任务：判断每个区域对应哪些节点下标，输出 JSON 映射，供引擎自动替换。
我会在下一条消息里给你文档按顺序编号的全部文字节点（编号=节点下标，格式「下标|文字」）。

输出 JSON 结构（字段严格一致；简历里没有的区域直接省略整个键）：
{{
 "fields": {{
   "name": [下标...], "ethnic": [...], "phone": [...], "email": [...],
   "birth": [...], "hukou": [...], "school": [...], "degree": [...],
   "major": [...], "courses": [...]
 }},
 "edu_line": {{"nodes": [...], "format": "{{period}}<按原样保留间隔>{{school}}"}},
 "slots": {{
   "objective": {{"nodes": [[a, b], ...]}},
   "proj_title": {{"nodes": [...]}},
   "proj_period": {{"nodes": [...]}},
   "modules": {{"nodes": [[...], [...]], "count": 每套节点数}},
   "tech": {{"nodes": [...], "prefix": "技术栈："}},
   "skills": {{"nodes": [[...], [...]], "count": 每套节点数}},
   "github_url": {{"nodes": [...], "label_nodes": [...]}},
   "demo_url": {{"nodes": [...], "label_nodes": [...]}}
 }},
 "cert_text": {{"nodes": [...], "bullet_nodes": [...], "label_nodes": [...], "clear_nodes": [[...]]}}
}}

各区域含义与判断规则：
1. fields 是基本信息：name=姓名的值、phone=电话的值、email=邮箱的值、birth=出生年月、
   hukou=户籍、school=毕业院校的值、degree=学历的值、major=专业的值、courses=主修课程的值。
   标签和值拆在不同节点时（如「姓 名」一个节点、「：张三」另一个节点），只列【值节点】；
   标签和值在同一个节点时（如「姓名：张三」），直接列该节点。
2. objective=求职意向；proj_title=项目标题行（项目名｜角色｜）；proj_period=项目时间；
   modules=项目要点（每条要点一个节点）；tech=技术栈行；skills=专业技能条目（每条一个节点）；
   github_url/demo_url=项目链接的 URL 节点；cert_text=证书行。
3. 很多 WPS 模板简历的每个内容存了两份、文字几乎相同（一份在形状里、一份在绘图里）。
   发现成对重复的内容，两套下标都要列（如 objective 写成 [[52,53],[54,55]]）；
   只出现一次的内容就列一套。modules/skills 的多套节点数量必须一致。
4. 宁缺勿错：拿不准的节点宁可不映射（省略），引擎会保留原样；
   严禁把同一个下标列进两个不同区域。
5. 细节判定：
   - modules/skills 只列【整句条目】的节点（通常以 ● 开头）；孤立的「●」符号节点、
     空节点、纯分隔符（如「/」）都不算条目，不要列进去。
   - cert_text 只列【证书正文】所在的节点（每套一个）；同一行里拆出来的
     其他碎片节点（如「证书：」标签、「/」分隔符、另一段正文）列入 clear_nodes。
6. 只输出 JSON，不要 markdown 包裹，不要任何解释。
"""


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------

def analyze_resume(path, original_name=""):
    """
    完整分析：档案草稿 + 专属模板包。
    返回 {"profile_draft": dict, "template": {"status","name","message"}}
    """
    lines = _extract_lines(path)
    draft = extract_profile(lines)

    tpl = {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
           "message": "未提供 docx，使用内置模板"}
    if path.lower().endswith(".docx"):
        try:
            tpl = build_user_template(path, original_name)
        except Exception as e:
            tpl = {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
                   "message": f"模板分析失败（{type(e).__name__}: {e}），使用内置模板"}
    return {"profile_draft": draft, "template": tpl}


def extract_profile(lines):
    """提示词一：简历原文行 → 档案草稿。LLM 失败返回手填骨架（全空）。"""
    text = "\n".join(lines)[:8000]
    if not text.strip():
        return _empty_draft()
    content, _, _ = call_model(
        messages=[{"role": "system", "content": _EXTRACT_SYSTEM},
                  {"role": "user", "content": "简历原文：\n\n" + text}],
        max_tokens=_EXTRACT_MAX_TOKENS, temperature=0.2, no_think=True)
    data = _parse_json(content)
    if not isinstance(data, dict):
        return _empty_draft()
    return normalize_profile(data)


def build_user_template(docx_path, original_name=""):
    """
    提示词二：dump 节点 → LLM 出映射 → 校验修复 → 落成模板包目录。
    返回 {"status": "ok"|"fallback", "name": 目录名, "message": ...}
    """
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(docx_path)
    ts = list(doc.element.body.iter(qn("w:t")))
    dump = "\n".join(f"{i}|{(t.text or '').strip()}"
                     for i, t in enumerate(ts) if (t.text or "").strip())
    if not dump.strip():
        return {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
                "message": "文档里没有文字，无法当模板"}

    content, _, _ = call_model(
        messages=[
            {"role": "system", "content": _MAP_SYSTEM},
            {"role": "user",
             "content": "文字节点清单：\n" + dump[:9000] + "\n\n请输出映射 JSON。"},
        ],
        max_tokens=_MAP_MAX_TOKENS, temperature=0.2, no_think=True)
    mapping = _parse_json(content)
    if not isinstance(mapping, dict):
        return {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
                "message": "模板结构分析失败，使用内置模板"}

    cfg, warnings, usable = _validate_mapping(mapping, len(ts))
    if not usable:
        msg = "；".join(warnings[-2:]) or "映射不可用"
        return {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
                "message": f"模板映射不可用（{msg}），使用内置模板",
                "missing": ["求职意向", "项目要点", "专业技能", "技术栈"]}

    # 关键区域逐个点名：哪个没识别到，生成时就会保持用户原文，必须让用户知道
    critical = [("objective", "求职意向"), ("modules", "项目要点"),
                ("skills", "专业技能"), ("tech", "技术栈")]
    missing = [label for key, label in critical if key not in cfg["slots"]]

    pack = f"user_{uuid.uuid4().hex[:8]}"
    d = os.path.join(TEMPLATES_DIR, pack)
    os.makedirs(d, exist_ok=True)
    shutil.copyfile(docx_path, os.path.join(d, "模板.docx"))
    cfg["display_name"] = original_name or "用户上传模板"
    cfg["docx"] = "模板.docx"
    cfg["min_nodes"] = len(ts)
    with open(os.path.join(d, "template.json"), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    msg = "已按你上传的简历版式建立专属模板"
    if missing:
        msg += ("；⚠「" + "、".join(missing) + "」没识别出来，"
                "生成时这些区域会保持你简历原文不动")
    if warnings:
        msg += "（" + "；".join(warnings[-2:]) + "）"
    return {"status": "ok", "name": pack, "message": msg, "missing": missing}


def normalize_profile(d):
    """把 LLM 输出/前端提交的档案整理成合法结构（缺键补默认，类型矫正）。"""
    p = d if isinstance(d, dict) else {}
    p["basic"] = _dict(p.get("basic"))
    p["education"] = _dict(p.get("education"))
    p["objective_default"] = str(p.get("objective_default") or "").strip()
    p["objective_candidates"] = _strlist(p.get("objective_candidates"))
    projects = p.get("projects") or []
    clean_projects = []
    for pr in projects:
        if not isinstance(pr, dict):
            continue
        pr = _dict(pr)
        pr["links"] = _dict(pr.get("links"))
        mods = []
        for m in (pr.get("modules") or []):
            if isinstance(m, dict) and str(m.get("name") or "").strip():
                mods.append({"name": str(m["name"]).strip(),
                             "core": bool(m.get("core")),
                             "default": str(m.get("default") or "").strip(),
                             "angles": _dict(m.get("angles"))})
        pr["modules"] = mods
        pr["tech_stack"] = _strlist(pr.get("tech_stack"))
        if str(pr.get("title") or "").strip():
            clean_projects.append(pr)
    p["projects"] = clean_projects
    skills = {}
    for g, items in (_dict(p.get("skills"))).items():
        arr = _strlist(items)
        if arr:
            skills[str(g).strip() or "技能"] = arr
    p["skills"] = skills
    p["certificates"] = _strlist(p.get("certificates"))
    rules = _dict(p.get("rules"))
    rules["style_rules"] = _strlist(rules.get("style_rules"))
    rules["tag_allow"] = _strlist(rules.get("tag_allow"))
    rules["tag_forbid"] = _strlist(rules.get("tag_forbid"))
    p["rules"] = rules
    scoring = _dict(p.get("scoring"))
    scoring["facts"] = _strlist(scoring.get("facts"))
    p["scoring"] = scoring
    p.setdefault("template", None)
    return p


# ---------------------------------------------------------------------------
# 映射校验：下标越界/一址多写/套内数量不一致都要修，修不动的整体回退
# ---------------------------------------------------------------------------

def _validate_mapping(mapping, n_nodes):
    warnings = []
    clean = {"slots": {}, "fields": {}, "edu_line": None}

    def ok_idx(i):
        return isinstance(i, int) and 0 <= i < n_nodes

    # 先把所有 entries 收集成 (键路径, 下标列表)，查冲突。
    # fields 兼容两种形态：扁平列表 [11,18] 或 {"nodes":[11,18]}（LLM 偶尔会包一层）
    entries = []          # (owner_key, [idx...])
    norm_fields = {}
    for k, v in (mapping.get("fields") or {}).items():
        if isinstance(v, dict) and isinstance(v.get("nodes"), list):
            v = v["nodes"]
        if isinstance(v, list) and v:
            norm_fields[k] = v
            entries.append((f"fields.{k}", v))
    slots = mapping.get("slots") or {}
    for k, v in slots.items():
        if not isinstance(v, dict):
            continue
        for kk in ("nodes", "label_nodes", "bullet_nodes"):
            if isinstance(v.get(kk), list):
                entries.append((f"slots.{k}.{kk}", v[kk]))
        for kk, vv in (v.get("clear_nodes") or {}).items() if isinstance(v.get("clear_nodes"), dict) else []:
            entries.append((f"slots.{k}.{kk}", vv))
    cert = mapping.get("cert_text") or {}
    if isinstance(cert, dict):
        for kk in ("nodes", "label_nodes", "bullet_nodes"):
            if isinstance(cert.get(kk), list):
                entries.append((f"cert_text.{kk}", cert[kk]))

    # 展开成 idx → 首次出现的 owner；重复出现的记冲突
    owner = {}
    conflict = set()
    for key, idxs in entries:
        for flat in _flatten(idxs):
            if not ok_idx(flat):
                continue
            if flat in owner and owner[flat] != key:
                conflict.add(flat)
            else:
                owner[flat] = key

    def scrub(idxs, own_key):
        """剔除越界和被其他区域抢占的下标；本区域先注册的保留（先到先得）。"""
        out = []
        for flat in _flatten(idxs):
            if ok_idx(flat) and owner.get(flat) in (None, own_key):
                out.append(flat)
        return out

    # fields（统一用规范化后的 norm_fields）
    for k, v in norm_fields.items():
        got = scrub(v, f"fields.{k}")
        if got:
            clean["fields"][k] = {"nodes": got, "mode": "after_colon"}
        else:
            warnings.append(f"fields.{k} 下标无效，已忽略")

    # slots
    for k, v in slots.items():
        if not isinstance(v, dict):
            continue
        entry = {}
        nodes = v.get("nodes")
        if isinstance(nodes, list) and nodes:
            # 嵌套列表（objective 的 [a,b] 对 / modules、skills 的成套下标）
            if all(isinstance(x, list) for x in nodes):
                sets = [scrub(x, f"slots.{k}.nodes") for x in nodes]
                sets = [s for s in sets if s]
                if k in ("modules", "skills"):
                    lens = {len(s) for s in sets}
                    if len(lens) > 1:
                        warnings.append(f"slots.{k} 各套槽数不一致，已忽略")
                        continue
                    if not sets:
                        warnings.append(f"slots.{k} 下标无效，已忽略")
                        continue
                    entry["nodes"] = sets
                    entry["count"] = len(sets[0])
                else:
                    entry["nodes"] = sets
                    if not sets:
                        warnings.append(f"slots.{k} 下标无效，已忽略")
                        continue
            else:
                got = scrub(nodes, f"slots.{k}.nodes")
                if not got:
                    warnings.append(f"slots.{k} 下标无效，已忽略")
                    continue
                entry["nodes"] = got
        else:
            warnings.append(f"slots.{k} 缺 nodes，已忽略")
            continue
        for kk in ("label_nodes", "bullet_nodes"):
            if isinstance(v.get(kk), list):
                got = scrub(v[kk], f"slots.{k}.{kk}")
                if got:
                    entry[kk] = got
        if isinstance(v.get("prefix"), str):
            entry["prefix"] = v["prefix"]
        clean["slots"][k] = entry

    # cert_text
    if isinstance(cert, dict) and cert.get("nodes"):
        got = scrub(cert["nodes"], "cert_text.nodes")
        if got:
            entry = {"nodes": got}
            for kk in ("label_nodes", "bullet_nodes"):
                if isinstance(cert.get(kk), list):
                    g = scrub(cert[kk], f"cert_text.{kk}")
                    if g:
                        entry[kk] = g
            cl = cert.get("clear_nodes")
            if isinstance(cl, list):
                g = [scrub(x, "cert_text.clear") if isinstance(x, list) else [] for x in cl]
                g = [x for x in g if x]
                if g:
                    entry["clear_nodes"] = g
            clean["slots"]["cert_text"] = entry

    # edu_line
    el = mapping.get("edu_line")
    if isinstance(el, dict) and isinstance(el.get("nodes"), list):
        got = scrub(el["nodes"], "edu_line")
        fmt = el.get("format") or "{period}          {school}"
        if got and "{period}" in fmt and "{school}" in fmt:
            clean["edu_line"] = {"nodes": got, "format": fmt}

    # 可用性判断：至少映射到 姓名 或 求职意向 或 要点区，否则当废
    usable = bool(clean["fields"].get("name") or clean["slots"].get("objective")
                  or clean["slots"].get("modules"))
    if not usable:
        warnings.append("有效映射太少")
    return clean, warnings, usable


def _flatten(x):
    """把任意嵌套的下标列表展平成 int 列表。"""
    if isinstance(x, int):
        return [x]
    if isinstance(x, list):
        out = []
        for i in x:
            out.extend(_flatten(i))
        return out
    return []


def _parse_json(text):
    if not text:
        return None
    t = text.strip()
    t = re.sub(r"^```(?:json)?\s*", "", t)
    t = re.sub(r"```\s*$", "", t)
    try:
        return json.loads(t)
    except Exception:
        pass
    a, b = t.find("{"), t.rfind("}")
    if a >= 0 and b > a:
        try:
            return json.loads(t[a:b + 1])
        except Exception:
            return None
    return None


def _empty_draft():
    d = normalize_profile({})
    d["template"] = template_export.DEFAULT_TEMPLATE
    return d


def _dict(x):
    return x if isinstance(x, dict) else {}


def _strlist(x):
    if not isinstance(x, list):
        return []
    return [str(i).strip() for i in x if str(i).strip()]


if __name__ == "__main__":
    # 自测：拿内置模板当"用户上传"跑一遍（它的正确答案就是 classic_blue 的 template.json）
    tpl = build_user_template(
        os.path.join(TEMPLATES_DIR, "classic_blue", "模板.docx"), "自测")
    print(json.dumps(tpl, ensure_ascii=False, indent=2))
    if tpl["status"] == "ok":
        with open(os.path.join(TEMPLATES_DIR, tpl["name"], "template.json"),
                  encoding="utf-8") as f:
            got = json.load(f)
        ref = json.load(open(os.path.join(TEMPLATES_DIR, "classic_blue",
                                          "template.json"), encoding="utf-8"))
        print("\n--- 与人工核对版对比 ---")
        print("LLM fields 键:", sorted(got.get("fields", {}).keys()))
        print("参考 fields 键:", sorted(ref.get("fields", {}).keys()))
        print("LLM slots  键:", sorted(got.get("slots", {}).keys()))
        print("参考 slots  键:", sorted(ref.get("slots", {}).keys()))
