# -*- coding: utf-8 -*-
r"""
模板替换导出（通用版，2026-09-15 M1 重构）：
拿 docx 模板当版式，只替换文字节点，形状/图片/版式原封不动。

模板包结构（一个目录 = 一个模板）：
  assets/templates/<模板名>/
    ├── 模板.docx        版式原稿
    └── template.json    节点索引 + 槽位配置（换模板 = 重新 dump 一份 json）

为什么用"替换文字"而不是 python-docx 从零画：
  原简历是 WPS 模板做的，含 23 个 v:shape + 20 个 w:drawing + 26 个文本框 + 1 张图片，
  从零造不出这些。唯一能"样子 100% 一样"的办法就是：打开原文件 → 只改文字 → 保存。

⚠️ 每个内容在 VML 形状和 DrawingML 绘图里各存一套，索引配置里两套都必须改，
   否则 Word 渲染哪一套取决于兼容模式，会出现前后不一致。
⚠️ 换新模板必须重新 dump 节点索引（tests/ 里有 dump 思路；模板.json 的 min_nodes 会兜底校验）。
"""

import json
import os

from docx import Document
from docx.oxml.ns import qn

import paths

DATA_ROOT = paths.data_root()
TEMPLATES_DIR = paths.templates_dir()
DEFAULT_TEMPLATE = "classic_blue"

MODULE_SLOTS_DEFAULT = 6
SKILL_SLOTS_DEFAULT = 5

VML_NS = "urn:schemas-microsoft-com:vml"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

W_HYPERLINK = qn("w:hyperlink")
R_ID = qn("r:id")


def list_templates():
    """列出可用的模板包名。"""
    if not os.path.isdir(TEMPLATES_DIR):
        return []
    out = []
    for d in sorted(os.listdir(TEMPLATES_DIR)):
        if os.path.isfile(os.path.join(TEMPLATES_DIR, d, "template.json")):
            out.append(d)
    return out


def load_pack(template_dir=None):
    """读模板包：返回 (docx 路径, 配置 dict)。"""
    name = template_dir or DEFAULT_TEMPLATE
    d = os.path.join(TEMPLATES_DIR, name)
    cfg_path = os.path.join(d, "template.json")
    if not os.path.isfile(cfg_path):
        raise RuntimeError(f"模板包不存在：{d}（需要有 template.json）")
    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)
    return os.path.join(d, cfg.get("docx", "模板.docx")), cfg


def _after_colon(node, value):
    """保留节点里"："前面的标签部分（含原始空格），只换后面的值。
    节点里没有冒号就整节点替换（用户上传的模板标签写法千奇百怪，兜底）。"""
    old = node.text or ""
    if "：" in old:
        node.text = old.rsplit("：", 1)[0] + "：" + value
    else:
        node.text = value


def _set(ts, idx, text):
    ts[idx].text = text


def export_from_result(result, out_path, profile=None, template_dir=None):
    """
    对外主入口：吃 generate.customize_resume() 的 result + 用户 profile，出 docx。

    result 用到：objective / project_title / project_role / project_points /
                 tech_stack / skill_lines / project_links
    profile 用到：basic（姓名/民族/电话/邮箱/出生/户籍/学历）、
                 education（起止/院校/专业/课程）、projects[0]（时间/链接兜底）、
                 certificates（证书行，可为空）
    """
    profile = profile or {}
    docx_path, cfg = load_pack(template_dir)
    doc = Document(docx_path)
    ts = list(doc.element.body.iter(qn("w:t")))
    min_nodes = cfg.get("min_nodes", 0)
    if len(ts) < min_nodes:
        raise RuntimeError(
            "模板结构变了：期望至少 %d 个 w:t 节点，实际 %d 个。"
            "请重新 dump 索引再改 template.json。" % (min_nodes, len(ts)))

    slots = cfg.get("slots", {})
    fields = cfg.get("fields", {})

    # ---------- 固定字段（全部来自 profile） ----------
    # 用户上传的模板不一定每个区域都有（比如没有民族/户籍），配置里缺哪个就跳过哪个
    basic = profile.get("basic") or {}
    edu = profile.get("education") or {}
    field_values = {
        "name": basic.get("name"),
        "ethnic": basic.get("ethnic"),
        "phone": basic.get("phone"),
        "email": basic.get("email"),
        "birth": basic.get("birth"),
        "hukou": basic.get("hukou"),
        "school": edu.get("school"),
        "degree": basic.get("degree"),
        "major": edu.get("major"),
        "courses": edu.get("courses"),
    }
    for key, conf in fields.items():
        v = field_values.get(key)
        if not v:
            continue
        for idx in conf.get("nodes", []):
            if conf.get("mode") == "after_colon":
                _after_colon(ts[idx], str(v))
            else:
                _set(ts, idx, str(v))
        for grp in conf.get("clear_nodes", []):
            for idx in grp:
                _set(ts, idx, "")

    # 教育行：起止时间 + 院校（一条）
    el = cfg.get("edu_line")
    if el and edu.get("period") and edu.get("school"):
        line = el.get("format", "{period}          {school}").format(
            period=str(edu["period"]).replace(" ", ""),
            school=edu["school"])
        for idx in el.get("nodes", []):
            _set(ts, idx, line)

    # ---------- 内容槽位（来自模型 result），缺配置的区域跳过 ----------
    # objective：节点可能被拆成多段（[a,b] 里 a 写值、b 及以后清空），也可能就一个节点
    conf = slots.get("objective")
    objective = (result.get("objective") or "").strip()
    if conf and objective:
        for pair in conf["nodes"]:
            _set(ts, pair[0], objective)
            for extra in pair[1:]:
                _set(ts, extra, "")

    conf = slots.get("proj_title")
    title = (result.get("project_title") or "").strip()
    role = (result.get("project_role") or "").strip()
    if conf and title:
        title_line = f"{title} ｜ {role} ｜ " if role else f"{title} ｜ "
        for idx in conf["nodes"]:
            _set(ts, idx, title_line)

    # 项目时间：优先 profile 里的项目时间
    conf = slots.get("proj_period")
    period = ((profile.get("projects") or [{}])[0].get("period") or "").strip()
    if conf and period:
        for idx in conf["nodes"]:
            _set(ts, idx, period)

    conf = slots.get("modules")
    modules = result.get("project_points")
    if conf and modules is not None:
        count = conf.get("count", MODULE_SLOTS_DEFAULT)
        for idx_list in conf["nodes"]:
            for k in range(count):
                _set(ts, idx_list[k], modules[k] if k < len(modules) else "")

    conf = slots.get("tech")
    tech = result.get("tech_stack")
    if conf and tech:
        prefix = conf.get("prefix", "技术栈：")
        line = prefix + "、".join(tech)
        for idx in conf["nodes"]:
            _set(ts, idx, line)

    # 项目链接：result 优先，profile 兜底；没有就整行清空
    links = result.get("project_links") or \
        ((profile.get("projects") or [{}])[0].get("links") or {})
    for key in ("github", "demo"):
        conf = slots.get(key + "_url")
        if not conf:
            continue
        url = (links.get(key) or "").strip()
        for idx in conf.get("nodes", []):
            _set(ts, idx, url)
        if not url:
            for idx in conf.get("label_nodes", []):
                _set(ts, idx, "")

    conf = slots.get("skills")
    skills = result.get("skill_lines")
    if conf and skills is not None:
        count = conf.get("count", SKILL_SLOTS_DEFAULT)
        for idx_list in conf["nodes"]:
            for k in range(count):
                _set(ts, idx_list[k], skills[k] if k < len(skills) else "")

    # 证书行：来自 profile，可整行消失；模板没有证书区就跳过
    cert = slots.get("cert_text")
    if cert:
        certs = [c.strip() for c in (profile.get("certificates") or []) if c.strip()]
        for idx in cert.get("nodes", []):
            _set(ts, idx, " / ".join(certs))
        if not certs:
            for idx in (cert.get("bullet_nodes") or []) + (cert.get("label_nodes") or []):
                _set(ts, idx, "")
        for grp in cert.get("clear_nodes", []):
            for idx in grp:
                _set(ts, idx, "")

    # 内容框改顶部对齐（内容比原稿少时不再上下悬空）
    _top_align_content_boxes(doc, cfg.get("content_box_cx", []))

    # 超链接目标同步：显示文字换了，点击跳转的 URL 也要跟着换
    #（超链接地址存在 rels 里，不改这里会出现"文字是 A 的 GitHub、点了却去 B 主页"的错配）
    _update_hyperlinks(doc, links, ts, slots)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


def _hyperlink_rids(ts, node_idxs):
    """从文字节点下标向上找它所在的 w:hyperlink，取那个关系 id。
    URL 是纯文本（没被包成超链接）时返回空——此时不去猜别的链接。"""
    rids = []
    for i in node_idxs:
        if not isinstance(i, int) or not (0 <= i < len(ts)):
            continue
        el = ts[i].getparent()
        while el is not None:
            if el.tag == W_HYPERLINK:
                rid = el.get(R_ID)
                if rid:
                    rids.append(rid)
                break
            el = el.getparent()
    return rids


def _update_hyperlinks(doc, links, ts, slots):
    """把 GitHub / 演示地址【所在的那个超链接】的跳转目标换成用户自己的。

    按节点定位，不按 host 猜：早先的实现是遍历全部关系、凡 target 含 github.com
    就归 GitHub、其余 http 一律归演示地址。用户模板里常有博客、作品集、二维码
    之类不相干的超链接，那样会把它们静默改成演示地址；用户没填 GitHub 时，
    模板原本的 GitHub 关系也会被 elif 抢走。
    """
    if not isinstance(links, dict):
        return
    try:
        rels = doc.part.rels
    except Exception:
        return
    want = []
    for key, slot in (("github", "github_url"), ("demo", "demo_url")):
        url = str(links.get(key) or "").strip()
        if not url:
            continue
        nodes = (slots.get(slot) or {}).get("nodes") or []
        want += [(rid, url) for rid in _hyperlink_rids(ts, nodes)]
    for rid, url in want:
        try:
            rels[rid]._target = url
        except Exception:
            pass


def _top_align_content_boxes(doc, box_cx_list):
    """把内容框改成顶部对齐。DrawingML 与 VML 两套都改，保证生效。"""
    body = doc.element.body
    cx_set = set(int(x) for x in box_cx_list)
    changed = 0

    # --- DrawingML（Word 实际渲染用的就是这套）---
    for w in body.iter("{%s}wsp" % WPS_NS):
        ext = None
        for e in w.iter("{%s}ext" % A_NS):
            if e.get("cx") and e.get("cy"):
                ext = e
                break
        if ext is None:
            continue
        try:
            if int(ext.get("cx")) not in cx_set:
                continue
        except (TypeError, ValueError):
            continue
        for bp in w.iter():
            if bp.tag.endswith("}bodyPr"):
                if bp.get("anchor") in ("ctr", "center"):
                    bp.set("anchor", "t")
                    changed += 1

    # --- VML：保持两套一致（按 style 里的 width 匹配）---
    for sh in body.iter("{%s}shape" % VML_NS):
        st = sh.get("style") or ""
        if not any("width:%spt" % (cx / 12700) in st for cx in cx_set):
            continue
        if "v-text-anchor:middle" in st:
            sh.set("style", st.replace("v-text-anchor:middle",
                                       "v-text-anchor:top"))
            changed += 1

    return changed


if __name__ == "__main__":
    # 冒烟测试：假数据跑一遍，校验结构保住、替换生效（不依赖模型）
    import re
    import zipfile

    OUT = os.path.join(paths.output_dir(), "_模板替换测试.docx")
    fake_profile = {
        "basic": {"name": "测试用例", "ethnic": "汉", "phone": "13800000000",
                  "email": "test@example.com", "birth": "2000.01",
                  "hukou": "测试市", "degree": "本科"},
        "education": {"period": "2021.09 - 2025.06", "school": "测试大学",
                      "major": "测试专业（本科）", "courses": "课程A、课程B"},
        "projects": [{"period": "2025.01 至今",
                      "links": {"github": "https://github.com/test/demo",
                                "demo": "https://demo.example.com"}}],
        "certificates": ["证书甲", "证书乙"],
    }
    fake_result = {
        "objective": "测试开发 实习生",
        "project_title": "测试项目", "project_role": "全栈开发",
        "project_points": ["●要点一", "●要点二", "●要点三",
                           "●要点四", "●要点五", "●要点六"],
        "tech_stack": ["技术1", "技术2"],
        "skill_lines": ["●技能一", "●技能二", "●技能三", "●技能四", "●技能五"],
    }
    p = export_from_result(fake_result, OUT, profile=fake_profile)
    print("已生成:", p)

    z = zipfile.ZipFile(p)
    xml = z.read("word/document.xml").decode("utf-8", "replace")
    print("形状 v:shape      :", len(re.findall(r"<v:shape[\s>]", xml)))
    print("绘图 w:drawing    :", len(re.findall(r"<w:drawing[\s>]", xml)))
    print("文本框 txbxContent:", len(re.findall(r"<w:txbxContent>", xml)))
    print("图片 media        :", len([n for n in z.namelist() if "/media/" in n]))

    d = Document(p)
    ts = list(d.element.body.iter(qn("w:t")))
    print("\n姓名节点 11:", repr(ts[11].text), "| 18:", repr(ts[18].text))
    print("电话节点 15:", repr(ts[15].text))
    print("邮箱节点 16:", repr(ts[16].text))
    print("教育行 38  :", repr(ts[38].text))
    print("意向节点 52:", repr(ts[52].text))
    print("要点节点 59:", repr(ts[59].text))
    print("技能节点 90:", repr(ts[90].text))
    print("证书节点 97:", repr(ts[97].text))
    print("链接节点 68:", repr(ts[68].text))
