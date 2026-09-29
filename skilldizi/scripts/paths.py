# -*- coding: utf-8 -*-
"""
数据根目录的唯一入口。

为什么要这个文件：内核（模板替换 / 档案 / 映射校验）原先用
`os.path.dirname(__file__)` 自己算数据在哪，也就是"代码住在哪，数据就在哪"。
一旦内核被搬进 skill 的 scripts/ 目录，而 profiles/ 和 assets/ 留在别处，
这行算出来的路径就全错了——而且是**静默指到空目录**，表现成"没有档案""模板不存在"。

现在四个内核模块一律只问这里，不再自己猜。

解析优先级（从高到低）：
  1. 环境变量 RESUME_DATA_DIR          —— 临时指定，测试用
  2. paths.py 旁边的 data_root.txt     —— 一行绝对路径，机器/部署级配置，不入库
  3. 从本文件往上找第一个同时含 profiles/ 和 assets/ 的目录
                                       —— 让"代码搬到哪，都能找回自家数据"，
                                          也是内核搬进 scripts/ 后零配置的根据
  4. 本文件所在目录                     —— 兜底，等同搬迁前的老行为
"""

import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_ENV_VAR = "RESUME_DATA_DIR"
_POINTER = os.path.join(_HERE, "data_root.txt")
_MARKERS = ("profiles", os.path.join("assets", "templates"))


def _has_data(d):
    return all(os.path.isdir(os.path.join(d, m)) for m in _MARKERS)


def _from_pointer():
    try:
        with open(_POINTER, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    return os.path.abspath(line)
    except OSError:
        pass
    return None


def _search_upward():
    d = _HERE
    while True:
        if _has_data(d):
            return d
        parent = os.path.dirname(d)
        if parent == d:
            return None
        d = parent


def data_root():
    env = os.environ.get(_ENV_VAR, "").strip()
    if env:
        return os.path.abspath(env)
    pointed = _from_pointer()
    if pointed:
        return pointed
    return _search_upward() or _HERE


def templates_dir():
    return os.path.join(data_root(), "assets", "templates")


def profiles_dir():
    return os.path.join(data_root(), "profiles")


def output_dir():
    return os.path.join(data_root(), "output")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(errors="replace")
    print("paths.py 位置 :", _HERE)
    print("data_root     :", data_root())
    print("templates_dir :", templates_dir())
    print("profiles_dir  :", profiles_dir())
    print("output_dir    :", output_dir())
