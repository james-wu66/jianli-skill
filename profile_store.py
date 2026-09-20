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


def list_profiles():
    """全部档案文件名（不含 active.txt）。"""
    if not os.path.isdir(PROFILES_DIR):
        return []
    return sorted(f for f in os.listdir(PROFILES_DIR) if f.endswith(".json"))


def active_name():
    """当前激活的档案文件名。"""
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
    return names[0]


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
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
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
    """补齐缺失键，保证引擎各处 .get / 下标访问不炸。"""
    p.setdefault("basic", {})
    p.setdefault("education", {})
    p.setdefault("objective_default", "")
    p.setdefault("objective_candidates", [])
    p.setdefault("projects", [])
    p.setdefault("skills", {})
    p.setdefault("certificates", [])
    p.setdefault("rules", {})
    p.setdefault("scoring", {})
    p.setdefault("template", None)
    p["rules"].setdefault("style_rules", [])
    p["rules"].setdefault("tag_allow", [])
    p["rules"].setdefault("tag_forbid", [])
    p["scoring"].setdefault("facts", [])
    if p["projects"]:
        p["projects"][0].setdefault("links", {})
        p["projects"][0].setdefault("period", "")
        p["projects"][0].setdefault("tech_stack", [])


if __name__ == "__main__":
    # 自检
    print("档案列表:", list_profiles())
    print("激活档案:", active_name())
    p = load_profile()
    print("姓名:", (p or {}).get("basic", {}).get("name"))
    print("技能分组:", list((p or {}).get("skills", {}).keys()))
