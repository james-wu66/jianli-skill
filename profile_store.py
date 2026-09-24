# -*- coding: utf-8 -*-
"""
个人档案（profile）存取层，2026-09-15 M1 新增。

每个用户一份 JSON，放在 profiles/ 目录，文件名即档案名（如 示例-张三.json）。
active.txt 里记当前激活的档案文件名；没有 active.txt 时取第一个（按文件名排序）。

profile 的结构约定（详见 profiles/示例-张三.json 的 _说明）：
  basic / education / objective_default / objective_candidates /
  projects / skills / certificates / rules / scoring / template
缺失的键会补上安全默认值，保证引擎不会 KeyError。
"""

import json
import os

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILES_DIR = os.path.join(_APP_DIR, "profiles")
ACTIVE_FILE = os.path.join(PROFILES_DIR, "active.txt")

# 示例档案（给别人照抄结构用的假人），永远不算"当前用户的档案"
EXAMPLE_PREFIX = "示例"


def list_profiles():
    """全部档案文件名（不含 active.txt）。"""
    if not os.path.isdir(PROFILES_DIR):
        return []
    return sorted(f for f in os.listdir(PROFILES_DIR) if f.endswith(".json"))


def real_profiles(names=None):
    """排除示例档案后的真实用户档案名列表。"""
    names = names if names is not None else list_profiles()
    return [n for n in names if not n.startswith(EXAMPLE_PREFIX)]


def active_name():
    """当前激活的档案文件名；没建档时返回 None（不是示例档案）。"""
    names = list_profiles()
    if not names:
        return None
    try:
        with open(ACTIVE_FILE, encoding="utf-8") as f:
            name = f.read().strip()
        if name in names:
            return name
    except OSError:
        pass
    # 没有 active.txt（新电脑上的常态）：只有存在非示例档案时才自动选一个。
    # 早先这里是"取目录里第一个"，而新仓库 profiles/ 里唯一的就是示例-张三.json，
    # 于是新用户会被判定成"已建档"，第一屏直接拿张三的假资料当他本人，
    # 向导再也不弹——README 里"首次使用会自动进入向导"那条也就是假的。
    real = real_profiles(names)
    return real[0] if real else None


def set_active(name):
    """切换激活档案。名字必须已存在。"""
    if name not in list_profiles():
        raise ValueError(f"档案不存在：{name}")
    with open(ACTIVE_FILE, "w", encoding="utf-8") as f:
        f.write(name)


def load_profile(name=None):
    """
    读档案并补齐默认值。name 传文件名（如 "张三.json"）或 None（用激活档案）。
    没有任何档案时返回 None（调用方应引导用户先建档案）。
    """
    name = name or active_name()
    if not name:
        return None
    path = os.path.join(PROFILES_DIR, name)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"档案文件不存在：{path}")
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"档案 {name} 不是合法 JSON：第 {e.lineno} 行 {e.msg}") from e
    if not isinstance(data, dict):
        raise ValueError(f"档案 {name} 最外层必须是一个 JSON 对象")
    _ensure_defaults(data)
    data["_file"] = name          # 内部字段：来源文件名
    return data


def save_profile(name, data):
    """写档案（向导/编辑页用）。name 必须以 .json 结尾。"""
    if not name.endswith(".json"):
        name += ".json"
    os.makedirs(PROFILES_DIR, exist_ok=True)
    path = os.path.join(PROFILES_DIR, name)
    data.pop("_file", None)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return name


def _ensure_defaults(p):
    """补齐缺失键【并矫正类型】。

    档案 JSON 完全可以手写（README 就推荐这条路），而这条路不过
    wizard.normalize_profile，所以类型兜底只能在这里做。之前只做了
    setdefault——键存在但值是 null / 列表时照样往下传，最后在引擎深处
    以 AttributeError 的形式炸出来（rules: null 实测会崩）。
    """
    for key, blank in (("basic", {}), ("education", {}), ("skills", {}),
                       ("certificates", []), ("objective_candidates", []),
                       ("projects", [])):
        if not isinstance(p.get(key), type(blank)):
            p[key] = {} if blank == {} else []
    if not isinstance(p.get("objective_default"), str):
        p["objective_default"] = ""
    if not isinstance(p.get("template"), str) or not p.get("template"):
        p["template"] = None

    rules = p["rules"] if isinstance(p.get("rules"), dict) else {}
    p["rules"] = rules
    for k in ("style_rules", "tag_allow", "tag_forbid"):
        if not isinstance(rules.get(k), list):
            rules[k] = []

    scoring = p["scoring"] if isinstance(p.get("scoring"), dict) else {}
    p["scoring"] = scoring
    if not isinstance(scoring.get("facts"), list):
        scoring["facts"] = []

    # 每一段项目都补，不只是第一段
    for pr in p["projects"]:
        if not isinstance(pr, dict):
            continue
        if not isinstance(pr.get("links"), dict):
            pr["links"] = {}
        if not isinstance(pr.get("period"), str):
            pr["period"] = ""
        if not isinstance(pr.get("tech_stack"), list):
            pr["tech_stack"] = []
        mods = pr.get("modules") if isinstance(pr.get("modules"), list) else []
        pr["modules"] = mods
        for m in mods:
            if not isinstance(m, dict):
                continue
            for k in ("name", "default"):
                if not isinstance(m.get(k), str):
                    m[k] = "" if m.get(k) is None else str(m.get(k))
            if not isinstance(m.get("angles"), dict):
                m["angles"] = {}
            m["core"] = bool(m.get("core"))


if __name__ == "__main__":
    # 自检
    print("档案列表:", list_profiles())
    print("激活档案:", active_name())
    p = load_profile()
    print("姓名:", (p or {}).get("basic", {}).get("name"))
    print("技能分组:", list((p or {}).get("skills", {}).keys()))
