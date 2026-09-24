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

from docx.oxml.ns import qn

from llm_client import call_model
from resume_parser import _extract_lines
import template_export

APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = template_export.TEMPLATES_DIR

# 提取档案：输出约 1~3 千 token，12000 足够且不触发长思考
_EXTRACT_MAX_TOKENS = 12000
# 模板映射：输出 1~2 千 token
_MAP_MAX_TOKENS = 8000

# 向导生成的专属模板包名（app.py 的删除接口用同一条规则）
USER_TPL_RE = re.compile(r"^user_[0-9a-f]{8}$")

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
我会在下一条消息里给你文档按顺序编号的全部文字节点，每行格式「下标 | 结构 | 文字」。
结构标记：D=这一行在 Word 绘图形状里，V=在 VML 形状里，P=普通段落，
N=这一行是 Word 自动项目符号列表的条目（文本里看不到符号，但它确实是一条列表项）。
同一处内容常常在 D 和 V 里各存一份（两份文字几乎一模一样）。

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
   "modules": {{"nodes": [该区域全部下标，扁平列表]}},
   "tech": {{"nodes": [...], "prefix": "技术栈："}},
   "skills": {{"nodes": [该区域全部下标，扁平列表]}},
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
3. 同一处内容常常存了两份（一份在 VML 形状里、一份在绘图里，对应结构标记 V 和 D）。
   处理方式分两种：
   - objective：要按"同一句话的两个副本"分组给出，如 [[52,53],[54,55]]；
   - modules / skills：**直接把这个区域涉及的全部下标列成一个扁平列表就行**，
     比如 [59,60,61,62,63,64,80,81,82,83,84,85]。分组由引擎按下标所在容器自动完成，
     你【不要】自己按"每条一套"或"每份一套"去嵌套——那种分法引擎会判为映射不可信。
     单值区域（proj_title/proj_period/tech/github_url/demo_url）同样给扁平列表。
4. 宁缺勿错：拿不准的节点宁可不映射（省略），引擎会保留原样；
   严禁把同一个下标列进两个不同区域。
5. 细节判定：
   - modules（项目要点）和 skills（技能条目）就是这两块里"一条一句"的整句条目节点，
     【不要靠开头有没有 ● 来认】。分点符号可能是 ● / · / • / - / ►，也可能什么都没有
     （结构标记带 N 的就是 Word 自动项目符号，符号不在文字里）。
     正确判据是：同属一个区域、句式长度相近、连续排列的若干完整句子。
     区域靠上下文中「项目经历」「专业技能」这类小标题，以及它们所在同一个形状/同一容器
     来确定。孤立的符号节点、空节点、纯分隔符（如「/」）不算条目，不要列进去。
   - 同一个项目里"●技术栈：…"这种汇总行属于 tech，不要算进 modules。
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


CRITICAL_REGIONS = (("objective", "求职意向"), ("modules", "项目要点"),
                    ("skills", "专业技能"), ("tech", "技术栈"))

# 判断"这份简历到底有没有这个区域"用的标题词。区域名各家叫法不同
# （实测见过「个人技能」「技能证书」「职业方向」），所以按别名集合找，
# 而不是只认提示词里那一个标准名。
REGION_HEADING_WORDS = {
    "objective": ("求职意向", "意向岗位", "期望职位", "求职意向：", "应聘"),
    "modules": ("项目经历", "项目经验", "个人项目", "作品集", "项目背景"),
    "skills": ("专业技能", "个人技能", "技能特长", "技能清单", "技术能力", "掌握技能"),
    "tech": ("技术栈", "涉及技术", "技术选型", "使用技术"),
}


def _regions_flat(cfg):
    """把校验后的 cfg 摊平成 {区域: [下标...]}，给前端做可编辑的初值。"""
    out = {}
    for key, _label in CRITICAL_REGIONS:
        v = (cfg.get("slots") or {}).get(key) or {}
        out[key] = [i for i in _flatten(v.get("nodes")) if isinstance(i, int)]
    return out


def update_template(pack, regions):
    """用户在页面上手工纠正过区域归属后重新落盘。

    映射这件事以前全自动、没人核对，所以模型一猜错就静默失效：那块内容
    保持原稿不动，用户拿到的是一份"看着正常但没定制过"的简历。
    这里给他一个改的机会，校验规则与自动路径完全同一套（不另开一条）。
    """
    d = os.path.join(TEMPLATES_DIR, pack)
    cfg_path = os.path.join(d, "template.json")
    docx_path = os.path.join(d, "模板.docx")
    if not USER_TPL_RE.match(pack) or not os.path.isfile(cfg_path):
        raise ValueError("模板包不存在：" + str(pack))

    from docx import Document
    from docx.oxml.ns import qn as _qn
    ts = list(Document(docx_path).element.body.iter(_qn("w:t")))
    _dump, containers = _dump_nodes(ts)
    old = json.load(open(cfg_path, encoding="utf-8"))

    # 以现有映射为底，只覆盖用户改过的区域；其余（fields/cert/链接）原样保留
    mapping = {"fields": {k: v.get("nodes") for k, v in (old.get("fields") or {}).items()},
               "slots": {k: v for k, v in (old.get("slots") or {}).items()
                         if k != "cert_text"},
               "cert_text": (old.get("slots") or {}).get("cert_text") or {},
               "edu_line": old.get("edu_line")}
    for key, _label in CRITICAL_REGIONS:
        if key in (regions or {}):
            picked = [i for i in regions[key] if isinstance(i, int)]
            if picked:
                mapping["slots"][key] = {"nodes": picked}
            else:
                mapping["slots"].pop(key, None)

    cfg, warnings, usable = _validate_mapping(mapping, len(ts), containers)
    if not usable:
        raise ValueError("纠正后的映射不可用：" + "；".join(warnings[-2:]))
    cfg["display_name"] = old.get("display_name") or "用户上传模板"
    cfg["docx"] = "模板.docx"
    cfg["min_nodes"] = old.get("min_nodes") or len(ts)
    cfg["content_box_cx"] = _content_box_cx(ts, cfg)
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    missing = [label for key, label in CRITICAL_REGIONS if key not in cfg["slots"]]
    # 项目要点和技能区全空 = 这份简历没有任何内容会被改写。
    # 不报错（确实有人没有项目经历），但必须让前端能亮红警示。
    no_content = not any(k in cfg["slots"] for k in ("modules", "skills"))
    return {"status": "ok", "name": pack, "missing": missing, "no_content": no_content,
            "regions": _regions_flat(cfg),
            "counts": {k: len(v) for k, v in _regions_flat(cfg).items()},
            "warnings": warnings}


def build_user_template(docx_path, original_name=""):
    """
    提示词二：dump 节点 → LLM 出映射 → 校验修复 → 落成模板包目录。
    返回 {"status": "ok"|"fallback", "name": 目录名, "message": ...}
    """
    from docx import Document
    doc = Document(docx_path)
    ts = list(doc.element.body.iter(qn("w:t")))
    dump, containers = _dump_nodes(ts)
    if not dump.strip():
        return {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
                "message": "文档里没有文字，无法当模板"}

    content, _, _ = call_model(
        messages=[
            {"role": "system", "content": _MAP_SYSTEM},
            {"role": "user",
             "content": "文字节点清单（格式「下标|结构|文字」）：\n"
                        + dump[:16000] + "\n\n请输出映射 JSON。"},
        ],
        max_tokens=_MAP_MAX_TOKENS, temperature=0.2, no_think=True)
    mapping = _parse_json(content)
    if not isinstance(mapping, dict):
        return {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
                "message": "模板结构分析失败，使用内置模板"}

    cfg, warnings, usable = _validate_mapping(mapping, len(ts), containers)
    if not usable:
        msg = "；".join(warnings[-2:]) or "映射不可用"
        return {"status": "fallback", "name": template_export.DEFAULT_TEMPLATE,
                "message": f"模板映射不可用（{msg}），使用内置模板",
                "missing": [label for _k, label in CRITICAL_REGIONS]}

    # 关键区域逐个点名：哪个没识别到，生成时就会保持用户原文，必须让用户知道
    # 区分"简历里没这块"和"有但没认出来"：前者不该报警。
    # 实测两份简历的侧栏标题是「技能证书」「个人技能」，而提示词里的区域名
    # 是"专业技能"，模型按字面找不到就整块不映射——报"没识别出来"是假警告。
    doc_text = " ".join((t.text or "") for t in ts)
    absent, real_missing = [], []
    for key, label in CRITICAL_REGIONS:
        if key in cfg["slots"]:
            continue
        if any(w in doc_text for w in REGION_HEADING_WORDS.get(key, ())):
            real_missing.append(label)
        else:
            absent.append(label)
    missing = real_missing

    pack = f"user_{uuid.uuid4().hex[:8]}"
    d = os.path.join(TEMPLATES_DIR, pack)
    os.makedirs(d, exist_ok=True)
    shutil.copyfile(docx_path, os.path.join(d, "模板.docx"))
    cfg["display_name"] = original_name or "用户上传模板"
    cfg["docx"] = "模板.docx"
    cfg["min_nodes"] = len(ts)
    cfg["content_box_cx"] = _content_box_cx(ts, cfg)
    with open(os.path.join(d, "template.json"), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    msg = "已按你上传的简历版式建立专属模板"
    if missing:
        msg += ("；⚠「" + "、".join(missing) + "」没识别出来，"
                "生成时这些区域会保持你简历原文不动 —— 可在下面手工指定")
    if absent:
        msg += "；你这份简历没有「" + "、".join(absent) + "」这块内容，不会动它"
    if warnings:
        msg += "（" + "；".join(warnings[-2:]) + "）"
    regions = _regions_flat(cfg)
    return {"status": "ok", "name": pack, "message": msg, "missing": missing,
            "absent": absent, "regions": regions,
            "counts": {k: len(v) for k, v in regions.items()},
            "samples": {k: [((ts[i].text or "").strip())[:34] for i in v[:2]]
                        for k, v in regions.items()},
            "nodes": [{"i": i, "text": (t.text or "").strip()}
                      for i, t in enumerate(ts) if (t.text or "").strip()]}


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

def _content_node_idxs(cfg, n_nodes):
    """多行内容区（项目要点、专业技能）用到的文字节点下标。
    只有这两块的内容条数会变、需要顶部对齐；求职意向/技术栈/标题都是单行，
    它们的框保持原稿的对齐方式，基本信息和教育背景更不该动。"""
    out = set()

    def add(x):
        for i in _flatten(x):
            if 0 <= i < n_nodes:
                out.add(i)

    slots = cfg.get("slots") or {}
    for k in ("modules", "skills"):
        v = slots.get(k)
        if isinstance(v, dict):
            add(v.get("nodes"))
    return out


def _as_sets(nodes):
    """把 nodes 统一成"若干套"的形态 [[..], [..]]。
    模型只给一套扁平列表时也当单套处理——以前这种会走到扁平分支，
    产出 entry["nodes"] 是整数列表，而 template_export 按"套"消费
    （idx_list[k]），整数下标会 TypeError。"""
    if nodes and all(isinstance(x, list) for x in nodes):
        return nodes
    return [nodes]


# 只认这套下标结构以外的区域，nodes 必须压平成整数列表
NESTED_SLOTS = ("modules", "skills", "objective")

WPS_WSP_TAG = "{%s}wsp" % template_export.WPS_NS
VML_SHAPE_TAG = "{%s}shape" % template_export.VML_NS
WPS_WSP = WPS_WSP_TAG
A_EXT = "{%s}ext" % template_export.A_NS


def _dump_nodes(ts, max_len=80):
    """生成给模型看的节点清单，并返回每个节点所在容器的标记表。

    清单每行「下标 | 结构 | 文字」，结构标记：
      D 在 Word 绘图形状里   V 在 VML 形状里   P 普通段落
      N 该段带 Word 自动项目符号（文本里没有符号，但它是一条列表项）

    两件事都必需：
    1) 早先判据写成"条目通常以 ● 开头"，可分点符号五花八门
       （· / • / - / 自动编号），这类简历的"项目要点/专业技能"整块认不出来。
       给结构线索，模型才能按语义和层级判断。
    2) containers 表让【引擎自己】按容器分套。同一份内容在 V 和 D 里各存一份，
       这个事实让模型去理解并不稳定（实测把"每条要点"当成"一套"，
       10 套 x 2 条会把同一句话写进 20 个位置、盖掉简历其他区域）。
    """
    lines = []
    containers = {}
    for i, t in enumerate(ts):
        txt = (t.text or "").strip()
        if not txt:
            continue
        flags = []
        numpr = False
        el = t.getparent()
        seen_shape = False
        while el is not None:
            if not seen_shape:
                if el.tag == WPS_WSP_TAG:
                    flags.append("D")
                    seen_shape = True
                elif el.tag == VML_SHAPE_TAG:
                    flags.append("V")
                    seen_shape = True
            if el.tag == qn("w:p") and el.find(".//" + qn("w:numPr")) is not None:
                numpr = True
            el = el.getparent()
        if numpr:
            flags.append("N")
        containers[i] = ("D" if "D" in flags else "V" if "V" in flags else "P")
        if len(txt) > max_len:
            txt = txt[:max_len] + "…"
        lines.append("%d|%s|%s" % (i, "".join(flags) or "P", txt))
    return "\n".join(lines), containers


# 同一区域的下标按容器分套的顺序：绘图一份、形状一份、普通段落一份
CONTAINER_ORDER = ("D", "V", "P")


def _group_by_container(flat, containers):
    """把一个区域的所有下标按容器分成若干套（保持文档顺序，丢掉空组）。"""
    groups = []
    for tag in CONTAINER_ORDER:
        g = [i for i in flat if containers.get(i, "P") == tag]
        if g:
            groups.append(g)
    return groups


def _content_box_cx(ts, cfg):
    """取映射节点所在形状（wps:wsp）的宽度 cx，写进模板包配置。
    template_export 靠它把内容框改成顶部对齐；缺了这个键，用户专属模板
    就会出现"内容比原稿少、文字在框里上下悬空"的排版问题——内置模板有、
    用户上传的没有，正是这里漏的。"""
    cxs = set()
    for i in _content_node_idxs(cfg, len(ts)):
        el = ts[i].getparent()
        while el is not None:
            if el.tag == WPS_WSP:
                for e in el.iter(A_EXT):        # 和 _top_align 取同一个 ext
                    if e.get("cx") and e.get("cy"):
                        try:
                            cxs.add(int(e.get("cx")))
                        except (TypeError, ValueError):
                            pass
                        break
                break
            el = el.getparent()
    return sorted(cxs)


def _validate_mapping(mapping, n_nodes, containers=None):
    warnings = []
    clean = {"slots": {}, "fields": {}, "edu_line": None}

    def ok_idx(i):
        return isinstance(i, int) and 0 <= i < n_nodes

    # 先把所有 entries 收集成 (键路径, 下标列表)，查冲突。
    entries = []          # (owner_key, [idx...])

    # edu_line 排在最前登记：引擎是先写 fields 再写 edu_line，同一节点被两者
    # 列到时以 edu_line 为准（它是"时间 + 院校"合成行，信息量更大）。
    # 早先这里漏登记，README 承诺的"一址多写校验"对 edu_line 根本无效。
    el = mapping.get("edu_line")
    if isinstance(el, dict) and isinstance(el.get("nodes"), list):
        entries.append(("edu_line", el["nodes"]))

    # fields 兼容两种形态：扁平列表 [11,18] 或 {"nodes":[11,18], "mode":..., "clear_nodes":...}
    # （LLM 一般给扁平列表；mode/clear_nodes 只有人工微调过的模板包会有，必须原样带过去，
    #   否则模板包过一次校验就被改写成清一色 after_colon、把碎片清空节点丢掉）
    norm_fields = {}
    for k, v in (mapping.get("fields") or {}).items():
        rmode = rclear = None
        if isinstance(v, dict) and isinstance(v.get("nodes"), list):
            rmode, rclear = v.get("mode"), v.get("clear_nodes")
            v = v["nodes"]
        if isinstance(v, list) and v:
            norm_fields[k] = (v, rmode, rclear)
            entries.append((f"fields.{k}", v))
            if isinstance(rclear, list):
                for grp in rclear:
                    if isinstance(grp, list):
                        entries.append((f"fields.{k}.clear", grp))
    slots = mapping.get("slots") or {}
    cert = mapping.get("cert_text")
    if not isinstance(cert, dict):
        # 有人（和某些模型）会把 cert_text 塞进 slots 里——存储格式就是那样的，两种都收
        inner = slots.get("cert_text")
        cert = inner if isinstance(inner, dict) else {}
    for k, v in slots.items():
        if k == "cert_text" or not isinstance(v, dict):
            continue
        for kk in ("nodes", "label_nodes", "bullet_nodes"):
            if isinstance(v.get(kk), list):
                entries.append((f"slots.{k}.{kk}", v[kk]))
    if isinstance(cert, dict):
        for kk in ("nodes", "label_nodes", "bullet_nodes"):
            if isinstance(cert.get(kk), list):
                entries.append((f"cert_text.{kk}", cert[kk]))
        # clear_nodes 是"和正文同行的碎片标签/分隔符"，会被真的写空，
        # 所以它的下标也要进冲突表，否则可能和别的区域抢同一个节点
        cl = cert.get("clear_nodes")
        if isinstance(cl, list):
            for grp in cl:
                if isinstance(grp, list):
                    entries.append(("cert_text.clear", grp))

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
    for k, (v, rmode, rclear) in norm_fields.items():
        got = scrub(v, f"fields.{k}")
        if not got:
            warnings.append(f"fields.{k} 下标无效，已忽略")
            continue
        conf = {"nodes": got,
                "mode": rmode if isinstance(rmode, str) and rmode else "after_colon"}
        if isinstance(rclear, list):
            groups = [scrub(g, f"fields.{k}.clear") for g in rclear if isinstance(g, list)]
            groups = [g for g in groups if g]
            if groups:
                conf["clear_nodes"] = groups
        clean["fields"][k] = conf

    # slots：按区域的消费方式分三种结构，不能一律原样透传
    #   modules / skills -> 若干套下标 [[..],[..]] + count
    #   objective        -> 成对下标 [[a,b],...]（a 写值，同组其余清空）
    #   其余单值区域      -> 必须是扁平整数下标
    # 以前非 modules/skills 的区域会把嵌套结构照原样留着，而 template_export 里
    # proj_title / proj_period / tech / github_url / demo_url 都是
    # `for idx in nodes: _set(ts, idx, ...)`——模型按提示词"两套都列"给出
    # [[..],[..]] 时，idx 是个列表，导出直接 TypeError 崩掉（实测踩过）。
    for k, v in slots.items():
        if k == "cert_text" or not isinstance(v, dict):
            continue
        entry = {}
        nodes = v.get("nodes")
        if not isinstance(nodes, list) or not nodes:
            warnings.append(f"slots.{k} 缺 nodes，已忽略")
            continue
        if k not in NESTED_SLOTS:
            got = scrub(nodes, f"slots.{k}.nodes")
            if not got:
                warnings.append(f"slots.{k} 下标无效，已忽略")
                continue
            entry["nodes"] = got
        elif k == "objective":
            groups = [scrub(x, f"slots.{k}.nodes") for x in _as_sets(nodes)]
            groups = [g for g in groups if g]
            if not groups:
                warnings.append(f"slots.{k} 下标无效，已忽略")
                continue
            entry["nodes"] = groups
        else:
            if containers:
                # 模型只负责"哪些下标属于这个区域"，分成几套由引擎按容器定。
                # 让模型自己理解"同一份内容在形状/绘图里各存一份"并不稳定：
                # 实测它会把"每条要点"当成"一套"，10 套 x 2 条会把同一句话
                # 写进 20 个位置、盖掉简历其他区域——改错比不改更糟。
                flat = []
                for i in _flatten(nodes):
                    if ok_idx(i) and i not in flat:
                        flat.append(i)
                raw_sets = _group_by_container(flat, containers)
            else:
                raw_sets = _as_sets(nodes)
            sets = [scrub(x, f"slots.{k}.nodes") for x in raw_sets]
            sets = [s for s in sets if s]
            if not sets:
                warnings.append(f"slots.{k} 下标无效，已忽略")
                continue
            if not containers and len(sets) > 2:
                # 没有容器信息可依据时（人工编辑的模板包），只能沿用"最多两套"
                # 这个形状约束来挡结构错乱的映射。
                warnings.append("slots.%s 被列成 %d 套（应为 1~2 套），判为映射不可信，已忽略"
                                % (k, len(sets)))
                continue
            lens = [len(s) for s in sets]
            if len(set(lens)) > 1:
                # 两套条数对不齐时按较少的那套走。以前这里是整块丢弃，
                # 后果是"项目要点/专业技能"整个区域完全不参与定制——
                # 简历最核心的两块白丢，只剩一句警告。
                n = min(lens)
                warnings.append("slots.%s 两套条数为 %s，按最少的 %d 条处理，"
                                "多出的位置保持你原文不动"
                                % (k, "/".join(str(x) for x in lens), n))
                sets = [s[:n] for s in sets]
            entry["nodes"] = sets
            entry["count"] = len(sets[0])
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

    # edu_line（el 已在上面登记冲突表时取出）
    if isinstance(el, dict) and isinstance(el.get("nodes"), list):
        got = scrub(el["nodes"], "edu_line")
        fmt = el.get("format")
        if not isinstance(fmt, str) or not fmt.strip():
            fmt = "{period}          {school}"
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
