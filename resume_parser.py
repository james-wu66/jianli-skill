# -*- coding: utf-8 -*-
r"""
简历解析器：从 .docx 文件提取结构化内容。
支持两种简历：常规段落型 / 文本框型（WPS 模板做的，内容全在 txbxContent 里）。

⚠️ 两个陷阱：
  1. 正则 <w:t[^>]*> 会误匹配 <w:tbl>/<w:tc>/<w:tab/> 等 w:t 开头的标签，
     必须写成 <w:t(?:\s[^>]*?)?>。
  2. 文本框顺序 ≠ 阅读顺序（"项目经历"标题可能排在最后），
     所以不能用"最近标题"归属，要用关键词启发式重分类。
"""
import os
import re
import zipfile

_TAG_T = re.compile(r"<w:t(?:\s[^>]*?)?>(.*?)</w:t>", re.S)


def _unesc(s):
    return (s.replace("&amp;", "&").replace("&lt;", "<")
             .replace("&gt;", ">").replace("&quot;", '"')
             .replace("&#39;", "'"))


def _texts_in(fragment):
    """从 XML 片段提取 w:t 文本，按段落聚合。"""
    out = []
    for para in re.split(r"</w:p>", fragment):
        joined = "".join(_TAG_T.findall(para)).strip()
        if joined:
            out.append(_unesc(joined))
    return out


def _extract_lines(path):
    """读出全部文本行，去重。"""
    z = zipfile.ZipFile(path)
    xml = z.read("word/document.xml").decode("utf-8")

    tbxs = re.findall(r"<w:txbxContent>(.*?)</w:txbxContent>", xml, re.S)
    if tbxs:
        seen, lines = set(), []
        for tb in tbxs:
            for ln in _texts_in(tb):
                if ln not in seen:
                    seen.add(ln)
                    lines.append(ln)
        return lines

    from docx import Document
    doc = Document(path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for tb in doc.tables:
        for row in tb.rows:
            row_t = "　".join(c.text.strip() for c in row.cells if c.text.strip())
            if row_t:
                lines.append(row_t)
    return lines


# ---------- 固定字段正则 ----------
_FIELDS = {
    "name": r"姓\s*名[：:]\s*([^\s]{2,5})",
    "phone": r"电\s*话[：:]\s*([\d\-\s]{10,})\s",
    "email": r"邮\s*箱[：:]?\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+)",
    "birth": r"出生年月[：:]\s*([\d/\.年月\-]+)",
    "hukou": r"户\s*籍[：:]\s*([\u4e00-\u9fa5]+?)(?:[\s　]|$)",
    "school": r"毕业院校[：:]\s*(\S+)",
    "degree": r"学\s*历[：:]\s*(\S+)",
}

# 装饰性文本（模板 slogan），直接丢弃
_NOISE = ("细心从每一个小细节开始", "Personal resume", "个人简历", "Resume")


def _classify_lines(lines):
    """按关键词启发式把行归到区块，不依赖标题出现顺序。"""
    buckets = {"education": [], "objective": [], "projects": [], "skills": []}

    for ln in lines:
        s = ln.strip()
        if not s or any(n in s for n in _NOISE):
            continue
        # 纯区块标题
        if s in ("基本信息", "教育背景", "求职意向", "项目经历", "专业技能",
                 "实习经历", "获奖情况", "自我评价", "个人优势"):
            continue

        # ① 技能分组行：●前端：/● 后端：/●AI：这种"分组名：内容"开头
        #    必须先于"项目要点"判断，否则"● 前端工程化："会被误判成项目
        if re.match(r"^[●·]?\s*(前端|后端|AI/数据|AI|数据|工程化|语言基础|证书|技能)[：:]", s):
            buckets["skills"].append(s)
            continue

        # ② 项目行：● 要点 / 含"技术栈" / 含 GitHub / 形如"项目名 ｜ 角色 ｜ 时间"
        if (s.startswith("●") or "技术栈" in s or "GitHub" in s
                or re.match(r"^[\u4e00-\u9fa5A-Za-z0-9（）()·]+[　 ]+｜[　 ]+", s)):
            buckets["projects"].append(s)
            continue

        # 教育行：含年份区间 + 学校
        if re.search(r"\d{4}\s*[.\-–]\s*\d{4}", s) or "主修课程" in s:
            buckets["education"].append(s)
            continue

        # 求职意向：短行且含"实习"或"开发"
        if len(s) <= 30 and ("实习" in s or "开发" in s):
            buckets["objective"].append(s)
            continue

    return {k: "\n".join(v) for k, v in buckets.items()}


_FILL_RE = re.compile(r'fillcolor="#?([0-9A-Fa-f]{6})"', re.I)
_COLOR_RE = re.compile(r'<w:color[^>]*w:val="([0-9A-Fa-f]{6})"', re.I)


def _lum(h):
    """感知亮度 0~255"""
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _sat(h):
    """饱和度：max-min"""
    v = [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    return max(v) - min(v)


def _dist(a, b):
    return sum(abs(int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) for i in (0, 2, 4))


def extract_theme(path):
    """
    从简历 docx 里提取配色，让别人上传什么配色就用什么配色。

    取色来源（都算上，因为 WPS 模板的颜色可能只存在于其中一处）：
      - VML 形状填充 fillcolor="#xxxxxx"（标题底块常用）
      - run 字体颜色 <w:color w:val="XXXXXX">

    规则：
      accent  主色   = 出现最多的彩色
      accent2 强调色 = 第二多的、且和主色拉开距离的彩色
      black   正文色 = 出现最多的中性色（黑白灰）
    找不到彩色就返回 {}（调用方回退默认配色）。

    返回 dict，键与 export.DEFAULT_THEME 一致。
    """
    try:
        z = zipfile.ZipFile(path)
        xml = z.read("word/document.xml").decode("utf-8", "replace")
    except Exception:
        return {}

    from collections import Counter
    cnt = Counter()
    for h in _FILL_RE.findall(xml):
        cnt[h.upper()] += 2          # 形状填充权重高，通常就是主色
    for h in _COLOR_RE.findall(xml):
        cnt[h.upper()] += 1

    # 去掉纯白纯黑，剩下的按"彩色 / 中性"分堆
    pool = {h: n for h, n in cnt.items() if h not in ("FFFFFF", "000000")}
    chroma = sorted(((h, n) for h, n in pool.items() if _sat(h) >= 25),
                    key=lambda x: (-x[1], -_sat(x[0])))
    neutral = sorted(((h, n) for h, n in pool.items() if _sat(h) < 25),
                     key=lambda x: -x[1])

    if not chroma:
        return {}

    accent = chroma[0][0]
    accent2 = None
    for h, _ in chroma[1:]:
        if _dist(h, accent) > 90:
            accent2 = h
            break
    if accent2 is None:
        accent2 = accent

    black = neutral[0][0] if neutral else "414141"
    # 底块上的字色：底亮就用深字，底暗就用白字
    title_text = "414141" if _lum(accent) > 150 else "FFFFFF"
    gold_text = "414141" if _lum(accent2) > 150 else "FFFFFF"

    return {
        "accent": accent,
        "accent2": accent2,
        "title_fill": accent,
        "title_text": title_text,
        "gold_fill": accent2,
        "gold_text": gold_text,
        "line": accent,
        "black": black,
        "gray": "666666",
    }


def parse_resume_docx(path):
    """
    解析简历，返回:
      fields   固定字段 dict
      sections 区块文本 dict（education/objective/projects/skills）
      lines    全部原始行（给 LLM 当补充素材）
    """
    lines = _extract_lines(path)

    fields = {}
    for k, pat in _FIELDS.items():
        m = re.search(pat, "\n".join(lines))
        if m:
            fields[k] = m.group(1).strip()

    return {
        "fields": fields,
        "sections": _classify_lines(lines),
        "lines": lines,
    }


if __name__ == "__main__":
    import sys
    p = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "assets", "templates", "classic_blue", "模板.docx")
    d = parse_resume_docx(p)
    print("=== 固定字段 ===")
    for k, v in d["fields"].items():
        print(f"  {k:8} = {v}")
    print("\n=== 区块 ===")
    for k, v in d["sections"].items():
        print(f"\n--- {k} ({len(v)} 字符) ---")
        print(v[:350])
