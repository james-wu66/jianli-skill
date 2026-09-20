# -*- coding: utf-8 -*-
"""
核心生成模块：JD 分析 + 简历定制 + 匹配度打分

核心思路（换汤不换药）：
  输入 = 用户档案 profile（唯一事实源，profiles/*.json） + 岗位 JD 原文
  输出 = 重排顺序/筛选侧重后的定制版简历结构 + 匹配度打分
  LLM 只做"排序和改写"，不做"编造"。

强制约束分两层（2026-09-15 M1 通用化拆分）：
  全局硬规则（写死在引擎里，任何用户不可关）：
    G1 禁止新增 profile 里没有的技能/项目/实习/荣誉（原 R5）
    G2 禁止虚构实习、公司经历或工作内容（原 R6）
    G3 证书/获奖不写进技能清单，由系统从 profile.certificates 单独写入证书栏
  个人规则（来自 profile.rules，每人可不同，见 profiles/示例-张三.json）：
    表述红线（原 R1/R2/R4/R7）+ 技能标签白/黑名单（原 R3）
"""

import json
import re
import time

from llm_client import call_model
from profile_store import load_profile

# max_tokens 是"上限"不是"目标"——模型写完 JSON 就停（finish_reason=stop）。
#
# ⚠️ 实测记录（都是 token.sensenova.cn 上真跑过的）：
#   100000 → HTTP 400 被拒：接口硬限制 [1, 65536]
#    65536 → 模型思考超过 600 秒，读超时，请求失败
#    20000 → 原来够用（旧 prompt 正文约 1000 字）
#    8000  → 52~76 秒出结果，正文完整
#
# ⚠️ 2026-09-10 晚踩的坑：改成"6 条项目要点 + 6 条技能"的新 prompt 后，
#   20000 被 reasoning token 全部吃光 → 主模型返回**空串**（不报错）→
#   被判定为失败 → 转备胎 → 备胎也挂 → 用户看到"模型都调不通"。
#   所以新 prompt 的输出量更大，预算必须相应提高。
#   取 32000：比 20000 大（保证正文写得出来），又远低于 65536（避免思考失控超时）。
MAX_TOKENS = 32000

# 技术栈兜底补充顺序（按"通用性"排序；只作为筛选器，实际只会补
# 用户档案 tech_stack 里确实存在的项，绝不无中生有）
TECH_FILL_PRIORITY = [
    "RESTful API", "GitHub Actions CI/CD", "Docker", "腾讯云 CloudBase",
    "RAG / LLM 集成", "Python Flask", "MySQL / SQLite",
    "JavaScript (ES6+)", "HTML / CSS / JavaScript (ES6+)",
]

# ---------------------------------------------------------------------------
# 事实源序列化：只把"事实"喂给模型，不给它发挥空间
# ---------------------------------------------------------------------------

def _profile_to_text(profile):
    """把用户档案转成给 LLM 看的紧凑事实清单。"""
    lines = []
    p = profile
    b = p.get("basic") or {}
    lines.append("【基本信息】")
    lines.append(f"姓名：{b.get('name', '')}｜电话：{b.get('phone', '')}｜邮箱：{b.get('email', '')}")
    lines.append(f"出生：{b.get('birth', '')}｜籍贯：{b.get('hukou', '')}")
    if b.get("github") or b.get("demo"):
        lines.append(f"GitHub：{b.get('github', '')}｜项目演示：{b.get('demo', '')}")

    e = p.get("education") or {}
    lines.append("\n【教育背景】")
    lines.append(f"{e.get('period', '')} {e.get('school', '')} {e.get('major', '')}")
    if e.get("courses"):
        lines.append(f"主修课程：{e['courses']}")

    lines.append(f"\n【默认求职意向】{p.get('objective_default', '')}")
    if p.get("objective_candidates"):
        lines.append("可选意向： " + " / ".join(p["objective_candidates"]))

    lines.append("\n【项目经历】")
    for pr in p.get("projects", []):
        lines.append(f"- {pr.get('title', '')}｜{pr.get('role', '')}｜{pr.get('period', '')}")
        if pr.get("intro"):
            lines.append(f"  概述：{pr['intro']}")
        for m in pr.get("modules", []):
            tag = "【核心】" if m.get("core") else "  "
            lines.append(f"  {tag}{m.get('name', '')}")
            if m.get("default"):
                lines.append(f"      默认表述：{m['default']}")
            for ang_name, ang_text in (m.get("angles") or {}).items():
                lines.append(f"      [{ang_name}角度] {ang_text}")
        if pr.get("tech_stack"):
            lines.append("  技术栈：" + "、".join(pr["tech_stack"]))

    lines.append("\n【专业技能库】（按 JD 需要重排/筛选，禁止新增）")
    for group, items in (p.get("skills") or {}).items():
        lines.append(f"- {group}：" + "；".join(items))

    if p.get("certificates"):
        lines.append("\n【证书】（由系统单独写入证书栏，不要写进技能清单）")
        lines.append("、".join(p["certificates"]))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Prompt
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Prompt：通用骨架（任何人一样） + 个人规则（来自 profile.rules）
# ---------------------------------------------------------------------------

_SYSTEM_ROLE = """你是资深中文技术简历顾问，专门帮应届本科生把项目经历改写成
不同岗位都投得进的定制版简历。

用户的项目经历可能有一段或多段（含实习/课程设计/个人项目），你做的全部工作都是
"挑选与岗位最相关的段落 + 重排顺序 + 调整侧重点 + 改写措辞"，把它们组织进简历的
项目经历区（版式固定若干条要点）。你没有新增事实的权限。
"""

# 全局硬规则：任何用户都生效，不允许通过 profile 关闭（产品底线）
_SYSTEM_GLOBAL_RULES = """必须遵守的硬性规则（违反任何一条即为不合格输出）：
G1 不得新增用户事实清单里没有的技能、项目、实习经历、奖项、工具。
G2 不得虚构任何实习经历、公司经历或工作内容。
G3 不要把证书/获奖类条目写进技能清单——证书由系统单独写入简历的证书栏，写了会重复。
"""

# 旧引擎里写死的 R1~R7 已拆走：
#   R1/R2/R4/R7 → profile.rules.style_rules（每人可自定义的表述红线）
#   R3          → profile.rules.tag_allow / tag_forbid（技能标签白/黑名单）
#   R5/R6       → 上面的 G1/G2（全局硬规则）


def _rules_text(profile):
    """把个人规则拼成提示词段落。没有个人规则就返回空串。"""
    rules = profile.get("rules") or {}
    lines = []
    sr = [r.strip() for r in rules.get("style_rules", []) if r and r.strip()]
    if sr:
        lines.append("本用户的个人表述红线（同样必须遵守）：")
        for i, r in enumerate(sr, 1):
            lines.append(f"P{i} {r}")
    allow = [t.strip() for t in rules.get("tag_allow", []) if t and t.strip()]
    forbid = [t.strip() for t in rules.get("tag_forbid", []) if t and t.strip()]
    if allow or forbid:
        seg = "技能标签方向限制："
        if allow:
            seg += "只能打 " + "、".join(allow) + " 这类方向"
        if allow and forbid:
            seg += "；"
        if forbid:
            seg += "绝对不能出现 " + "、".join(forbid) + " 等标签"
        seg += "。"
        lines.append(seg)
    return ("\n".join(lines) + "\n\n") if lines else ""


def build_system_prompt(profile):
    """组装系统提示词：通用角色 + 全局硬规则 + 个人规则。"""
    return _SYSTEM_ROLE + "\n" + _SYSTEM_GLOBAL_RULES + "\n" + _rules_text(profile) \
        + "输出必须是合法 JSON（不要 markdown 代码块包裹，不要 JSON 外的任何文字）。"

_JSON_SCHEMA = '''{
  "match": {
    "score": 70,
    "recommend": "值得投/可以投/建议不投",
    "summary": "一句总评，30字内",
    "requirements": [
      {"item": "JD 要求点", "status": "matched/partial/missing", "note": "差距说明，20字内"}
    ]
  },
  "fit": "岗位方向，按 JD 实际判断（前端/后端/全栈/AI应用/数据/算法/测试/运维/其他）",
  "objective": "定制后的求职意向一句话，20字内",
  "project_title": "项目名称（取档案里最相关的项目）",
  "project_role": "本项目里对应的岗位角色，如 前端开发 / 后端开发 / 全栈开发",
  "project_points": [
    "●【核心】模块名：内容，45~90个汉字",
    "（固定 6 条，多余的不要）"
  ],
  "tech_stack": ["只放与该岗位方向相关的技术，4~6 项"],
  "skill_lines": [
    "●分组名：内容，40~80个汉字",
    "（固定 5 条，证书由系统单独写入，不要包含证书条目）"
  ],
  "rewrite_notes": "本次改了什么，25字内"
}'''


def _build_user_prompt(jd_text, focus_hint, extra_material="", profile=None):
    profile = profile or {}
    facts = _profile_to_text(profile)
    hint_line = (
        '用户额外说明的侧重点：' + focus_hint
    ) if focus_hint else '（用户未指定额外侧重，请根据 JD 自行判断）'

    # 用户当次上传的简历原文：作为补充事实素材，优先级低于上面那份结构化清单
    extra_block = ""
    if extra_material and extra_material.strip():
        extra_block = (
            "\n\n\n=== 用户本次上传的简历原文（补充素材）===\n"
            "以下是用户当次上传的简历提取出的文字。它和上面的事实清单可能有重叠，\n"
            "也可能包含清单里没有的细节。**只用它补充事实，不允许超出它和清单的范围编造。**\n"
            "如果两者冲突，以上面的事实清单为准。\n\n"
            + extra_material.strip()[:4000]
            + "\n=== 上传简历结束 ==="
        )

    return (
        "下面是用户唯一的真实简历事实清单（唯一信息源，禁止超出范围编造）：\n\n"
        + facts
        + extra_block
        + "\n\n\n=== 目标岗位 JD 原文 ===\n"
        + jd_text.strip()[:3500]
        + "\n=== JD 结束 ===\n\n"
        + hint_line
        + "\n\n请完成两件事：\n\n"
        "任务一：把上面的事实清单，针对这份 JD 重新组织成定制版简历。\n"
        "  - 分析 JD 的岗位要求属于哪个方向。**不限于**前端/后端/全栈/AI应用/数据，"
        "测试、运维、Python 开发等任何方向都按 JD 实际判断，填进 fit 字段。\n"
        "  - project_points【6 条】：每条写成「●模块名：内容」，核心的 2 条加【核心】前缀。\n"
        "    按与 JD 的相关度排序，最相关的排最前；优先用【与该岗位方向同名的那条角度】的表述。\n"
        "    长度跟事实清单里那句原话差不多就行，不要为凑字数扩写。\n"
        "  - skill_lines【5 条】：每条写成「●分组名：内容」，按与 JD 的相关度排序。\n"
        "    长度跟事实清单里那段差不多就行；不相关的分组改成相关角度，但不得编造。\n"
        "    ⚠️ 不要写证书/获奖这类条目 —— 证书由系统单独写入简历的证书栏，写了会重复。\n"
        "  - tech_stack（技术栈，极其重要）：从项目技术栈里【只挑与该岗位方向直接相关的 4~6 项】。\n"
        "    · 投后端 → 只留 Python/Flask、数据库 SQL、RESTful API、Docker、CI/CD 这类；\n"
        "      严禁出现 JavaScript、HTML/CSS、PWA、前端动画等前端项。\n"
        "    · 投前端 → 只留 HTML/CSS/JavaScript(ES6+)、SPA 架构、移动端响应式适配、前端动画这类；\n"
        "      严禁出现数据库优化、鉴权方案、容器化部署等后端项。\n"
        "    · 投 AI 应用 → 只留 LLM 集成、RAG、数据管道、Python 这类。\n"
        "    · 投全栈 → 前端、后端各留 2~3 项，两边都必须出现，不要只剩一边。\n"
        "    · 宁可少列，也绝对不要全列 —— 把 9 项全列出来，面试官会认为你没有专精方向。\n"
        "  - 求职意向：**JD 里写了明确职位名时，直接照抄 JD 的职位名**（按需补上\"实习生\"字样，20 字内），\n"
        "    禁止换成可选意向里的其他叫法；JD 没写职位名时，才从可选意向里选最贴的那个，20 字内。\n"
        "  - project_role：写本项目里对应的岗位角色，如「前端开发」「后端开发（Python）」。\n\n"
        "任务二：给出匹配度分析。\n"
        "  - 从 JD 里提取 4~8 个关键要求点（技术栈、能力、经验）。\n"
        "  - 逐条判断用户匹配情况：matched（完全匹配）/ partial（部分匹配）/ missing（缺失）。\n"
        "  - 给出总体匹配分 0~100 和建议值不值得投。\n\n"
        "匹配分评分纪律（必须严格执行，禁止虚高）：\n"
        + _scoring_text(profile) +
        "  - 凡是 JD 提到用户事实清单里确实没有的内容（实习经历、团队协作、企业级项目、"
        "特定业务经验、高级框架、高并发/微服务等），"
        "必须判 missing 并如实扣分，不要因为简历写得好就加分掩盖。\n"
        "  - 分数应反映真实拿到面试的概率，不是简历好不好看。\n"
        "    参考基准：4 条全 matched 通常 65~75 分；出现 1 条 partial 扣 5~10 分；"
        "出现 1 条 missing 扣 8~15 分。多数应届生对大厂岗位真实匹配分在 40~70 之间，"
        "打 85 分以上必须有充分理由。\n"
        "  - 宁可诚实偏低，也不要让用户拿着虚高分去投明显不适合的岗位浪费时间。\n\n"
        "注意：输出要精简。project_points 6 条、skill_lines 5 条，"
        "tech_stack 只放 4~6 项且与该岗位方向一致。\n\n"
        "严格按以下 JSON 结构输出：\n"
        + _JSON_SCHEMA
        + "\n\n现在开始，只输出 JSON。"
    )


def _scoring_text(profile):
    """评分纪律开头：把用户档案里的真实短板事实拼进去，防止打分虚高。"""
    facts = [f.strip() for f in (profile.get("scoring") or {}).get("facts", [])
             if f and f.strip()]
    if not facts:
        return ""
    return ("  - 用户的真实短板是客观存在的：" + "、".join(facts) + "。\n")


# ---------------------------------------------------------------------------
# 容错解析
# ---------------------------------------------------------------------------

def _extract_json(text, profile):
    """
    解析模型输出的 JSON。
    处理三类情况：完整 JSON / markdown 包裹 / 被 max_tokens 截断。
    """
    if not text:
        return None
    t = text.strip()
    t = re.sub(r"^```(?:json)?\s*", "", t)
    t = re.sub(r"```\s*$", "", t)

    try:
        obj = json.loads(t)
        return obj if isinstance(obj, dict) else None
    except Exception:
        pass

    # 截断恢复：从末尾往前找最后一个 '}'，尝试补 ] 和 }
    for cut in range(len(t) - 1, max(200, len(t) - 1200), -1):
        if t[cut] != "}":
            continue
        candidate = t[:cut + 1]
        for suffix in ("", "]", "}]", "}}", "}]", "}]]", "}}]", "}}}"):
            try:
                obj = json.loads(candidate + suffix)
                if isinstance(obj, dict):
                    return _fill(obj, profile)
            except Exception:
                continue
    return None


def _fill(obj, profile):
    """补齐模型未输出的可选字段，保证导出时不会 KeyError。"""
    projects = profile.get("projects") or [{}]
    pr = projects[0]
    if not obj.get("match"):
        obj["match"] = {
            "score": None, "recommend": "无法评估",
            "summary": "模型输出被截断，匹配分析未完成。",
            "requirements": [],
        }
    # 模板要求 6 条要点 / 5 条技能，缺了就用基础素材补
    if not obj.get("project_points"):
        obj["project_points"] = ["●" + m["name"] + "：" + m["default"]
                                 for m in pr.get("modules", [])]
    if not obj.get("skill_lines"):
        # 证书组不进技能清单（由系统单独写入证书栏），最多取 5 组
        groups = [(g, items) for g, items in (profile.get("skills") or {}).items()
                  if "证书" not in g][:5]
        obj["skill_lines"] = ["●" + g + "：" + "；".join(items)
                              for g, items in groups]
    obj.setdefault("rewrite_notes", "（模型输出被截断，说明缺失）")
    obj.setdefault("project_title", pr.get("title", ""))
    obj.setdefault("project_role", pr.get("role", ""))
    obj.setdefault("project_links", pr.get("links", {}))
    obj.setdefault("tech_stack", pr.get("tech_stack", []))
    # 技术栈不足 4 项时用项目真实技术栈补足 —— 不是编造，都是项目确实用过的，
    # 只是模型偶尔给太少（实测出现过只给 2 项，简历看着单薄）
    ts = list(obj.get("tech_stack") or [])
    if len(ts) < 4:
        # 按"通用性"补：RESTful / CI/CD 这类对任何岗位都合适，
        # 避免给前端岗补上 Python Flask、MySQL 这种后端专属项
        real = set(pr.get("tech_stack") or [])
        for t in TECH_FILL_PRIORITY:
            if t not in ts and t in real:
                ts.append(t)
            if len(ts) >= 4:
                break
        obj["tech_stack"] = ts
    obj.setdefault("fit", "全栈")
    obj.setdefault("objective", profile.get("objective_default", ""))
    return obj


def _fallback_result(error, profile):
    """LLM 全挂时的降级结果：不编造，直接返回基础简历 + 错误说明。"""
    projects = profile.get("projects") or [{}]
    pr = projects[0]
    # 档案里的模块 / 技能组直接拼成模板要的 6+5 条（证书组除外）
    points = ["●" + m["name"] + "：" + m.get("default", "")
              for m in pr.get("modules", [])]
    groups = [(g, items) for g, items in (profile.get("skills") or {}).items()
              if "证书" not in g][:5]
    skills = ["●" + g + "：" + "；".join(items) for g, items in groups]
    return {
        "fit": "全栈",
        "objective": profile.get("objective_default", ""),
        "project_title": pr.get("title", ""),
        "project_role": pr.get("role", ""),
        "project_points": points,
        "tech_stack": pr.get("tech_stack", []),
        "skill_lines": skills,
        "match": {
            "score": None,
            "recommend": "无法评估",
            "summary": f"模型调用失败（{error}），以下为未定制的原始简历内容。",
            "requirements": [],
        },
        "rewrite_notes": "本次未调用到模型，输出为基础简历原样。",
        "_fallback": True,
        "_error": error,
    }


# ---------------------------------------------------------------------------
# 对外主入口
# ---------------------------------------------------------------------------

def customize_resume(jd_text, focus_hint="", extra_material="", profile=None):
    """
    输入 JD（+可选重点 +可选上传简历原文）+ 用户档案，
    返回 (定制结果 dict, 元信息 dict)。
    元信息含 model_used / usage / elapsed / fallback。
    profile 为 None时自动加载当前激活档案。
    """
    if profile is None:
        profile = load_profile()
    if not profile:
        raise ValueError("没有个人档案（profiles/ 目录为空），请先完成首次配置")

    t0 = time.time()

    if not jd_text or not jd_text.strip():
        return _fallback_result("空 JD", profile), {"fallback": True, "error": "JD 为空"}

    content, usage, model_used = call_model(
        messages=[
            {"role": "system", "content": build_system_prompt(profile)},
            {"role": "user",
             "content": _build_user_prompt(jd_text, focus_hint,
                                           extra_material, profile)},
        ],
        max_tokens=MAX_TOKENS,
        temperature=0.45,
    )
    meta = {
        "model_used": model_used,
        "usage": usage,
        "elapsed": round(time.time() - t0, 1),
        "fallback": False,
        "used_upload": bool(extra_material and extra_material.strip()),
    }

    if not content:
        meta["fallback"] = True
        # 给用户看得懂、可操作的原因，而不是干巴巴一句"模型全部不可用"
        meta["error"] = ("模型暂时都调不通。多数情况是主模型撞了限流，"
                         "一般等 2~3 分钟会自动恢复 —— 稍后再点一次生成即可。")
        return _fallback_result("模型全部不可用", profile), meta

    data = _extract_json(content, profile)
    if not data:
        meta["fallback"] = True
        meta["error"] = "模型返回内容无法解析为 JSON"
        return _fallback_result("返回内容无法解析", profile), meta

    data.setdefault("project_points", [])
    data.setdefault("skill_lines", [])
    data.setdefault("tech_stack", [])
    data.setdefault("match", {})
    data.setdefault("rewrite_notes", "")
    return data, meta


if __name__ == "__main__":
    demo_jd = """【岗位职责】
1. 负责后端服务的设计与开发，参与性能优化
2. 参与数据库设计与接口联调

【任职要求】
1. 熟悉 Python Flask 或 Django
2. 有实际项目部署上线经验
3. 熟悉 Git 版本管理
4. 有团队协作经验者优"""

    result, meta = customize_resume(demo_jd)
    print(f"model: {meta['model_used']} | 耗时 {meta['elapsed']}s | fallback: {meta['fallback']}")
    print(f"fit: {result['fit']}")
    print(f"objective: {result['objective']}")
    print(f"\n项目要点（{len(result['project_points'])} 条）:")
    for x in result["project_points"]:
        print(f"  {x}")
    print(f"\n技能（{len(result['skill_lines'])} 条）:")
    for x in result["skill_lines"]:
        print(f"  {x}")
    print(f"\n技术栈: {'、'.join(result.get('tech_stack') or [])}")
    m = result["match"]
    print(f"\n匹配度: {m['score']} / {m['recommend']}")
    print(f"总评: {m['summary']}")
    print("\n要求逐条:")
    for r in m.get("requirements", []):
        print(f"  [{r['status']:7}] {r['item']}  —  {r['note']}")
    print(f"\n改写说明: {result['rewrite_notes']}")
