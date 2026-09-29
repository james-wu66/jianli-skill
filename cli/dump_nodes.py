# -*- coding: utf-8 -*-
"""
CLI：把一份简历 .docx 摊成给模型/Agent 看的文字节点清单。

    python cli/dump_nodes.py 某份简历.docx [--out 节点清单.txt]

输出每行「下标|结构|文字」，结构标记 D=绘图形状 V=VML形状 P=普通段落 N=该段带自动编号。

⚠️ 这里**故意不输出 containers 分套表**。同一份内容在 V 和 D 里各存一份，这个
分套由引擎在 apply_map 时按节点容器自己判断。交出去让模型理解并不稳定——实测它
会把"每条要点"当成"一套"，产出 10 套 × 2 条，引擎就把同一句话写进 20 个位置、
盖掉简历其他区域。
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx.oxml.ns import qn

import wizard


def main():
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="导出 docx 文字节点清单")
    ap.add_argument("docx")
    ap.add_argument("--out", default="", help="写进这个文件（UTF-8）；不填就打屏")
    ap.add_argument("--max-len", type=int, default=80, help="每条文字最多显示几个字")
    a = ap.parse_args()

    ts = list(Document(a.docx).element.body.iter(qn("w:t")))
    dump, _containers = wizard._dump_nodes(ts, a.max_len)
    listed = len(dump.splitlines()) if dump.strip() else 0
    head = "# 节点总数(含空) %d ｜ 有文字的 %d ｜ 格式 下标|结构|文字\n" % (len(ts), listed)

    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(head + dump + "\n")
        print('{"ok": true, "out": "%s", "n_nodes": %d, "listed": %d}'
              % (a.out.replace("\\", "/"), len(ts), listed))
    else:
        sys.stdout.write(head + dump + "\n")

    if listed == 0:
        print('{"ok": false, "error": "文档里没有文字，无法当模板"}')
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
