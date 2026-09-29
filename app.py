# -*- coding: utf-8 -*-
"""
简历生成器 Web 应用
运行：python app.py   →  http://127.0.0.1:5050

接口：
  GET  /              单页界面
  POST /api/generate  {jd, focus} → 定制结果 + docx 路径
  GET  /output/<f>    下载生成的 Word
"""

import json
import os
import re
import shutil
import sys
import tempfile
import threading
import time
import uuid

from flask import Flask, request, jsonify, send_file

APP_DIR = os.path.dirname(os.path.abspath(__file__))
# 内核（模板原地换字 / 档案 / 映射校验）住在 skilldizi/scripts，
# 网页版与 skill 版共用同一份代码，不复制第二份。
# RESUME_DATA_DIR 必须赶在 import 内核之前设好：paths.py 会向上找第一个同时含
# profiles/ 与 assets/ 的目录，内核从 skilldizi/scripts 出发会找到 skill 那份档案；
# 而按现在的约定两个入口各用各的数据目录（网页版=项目根）。
os.environ.setdefault("RESUME_DATA_DIR", APP_DIR)
sys.path.insert(0, os.path.join(APP_DIR, "skilldizi", "scripts"))

import template_export
from generate import customize_resume
from template_export import export_from_result
from resume_parser import parse_resume_docx
import profile_store
import wizard as wizard_mod

OUTPUT_DIR = os.path.join(APP_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 导出模板目录（assets/templates/<名>/，含 模板.docx + template.json）。
# 2026-09-15 M1 起由用户档案的 template 字段决定（默认 classic_blue），
# 环境变量 RESUME_TEMPLATE 仍可强制覆盖。
TEMPLATE_OVERRIDE = os.environ.get("RESUME_TEMPLATE")

MAX_UPLOAD = 8 * 1024 * 1024   # 8 MB 上限
MAX_TASKS = 60                 # 最多留 60 条任务记录，超了丢最早的

app = Flask(__name__, static_folder=OUTPUT_DIR, static_url_path="/output")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD

# ---------------------------------------------------------------------------
# 后台任务池
# 生成要 60~150 秒，不能让浏览器干等（等太久一刷新就 Failed to fetch）。
# 改成：POST 立刻返回 task_id，前端轮询 /api/status/<id>。
# ---------------------------------------------------------------------------
_TASKS = {}
_LOCK = threading.Lock()


def _set_task(tid, **kw):
    with _LOCK:
        _TASKS.setdefault(tid, {}).update(kw)


def _get_task(tid):
    with _LOCK:
        return dict(_TASKS.get(tid) or {})


def _new_task():
    with _LOCK:
        # 超量了就按创建顺序丢掉最早的一批
        if len(_TASKS) >= MAX_TASKS:
            for k in list(_TASKS.keys())[:len(_TASKS) - MAX_TASKS + 1]:
                _TASKS.pop(k, None)
        tid = uuid.uuid4().hex[:12]
        _TASKS[tid] = {"status": "queued", "created": time.time()}
        return tid


def _run_task(tid, jd, focus, extra_material, uploaded_name, profile):
    """后台线程：调模型 + 出 Word。"""
    _set_task(tid, status="running", started=time.time())
    t0 = time.time()
    # 本次真正要写入的模板包：先定下来，条数按它算，导出也用它。
    # 两边用同一个值，才不会出现"条数按 A 模板产出、内容往 B 模板里写"。
    use_template = TEMPLATE_OVERRIDE or profile.get("template")
    try:
        result, meta = customize_resume(jd, focus, extra_material, profile,
                                        template=use_template)
    except Exception as e:
        _set_task(tid, status="error",
                  error=f"生成过程出错：{type(e).__name__}: {e}")
        return

    if meta.get("fallback"):
        _set_task(tid, status="error",
                  error=meta.get("error") or "模型不可用，没能生成定制版",
                  meta=meta)
        return

    try:
        fit = (result.get("fit") or "全栈").strip()
        who = (profile.get("basic") or {}).get("name") or "简历"
        # 文件名带时间戳：同方向投多家不会互相覆盖（海投场景每次生成都是独立一份）
        stamp = time.strftime("%m%d_%H%M%S")
        fname = f"{who}简历_{fit}_{stamp}.docx"
        path = os.path.join(OUTPUT_DIR, fname)
        export_from_result(result, path, profile=profile,
                           template_dir=use_template)
    except Exception as e:
        _set_task(tid, status="error",
                  error=f"Word 导出失败：{type(e).__name__}: {e}")
        return

    meta["elapsed"] = round(time.time() - t0, 1)
    meta["file"] = os.path.basename(path)
    meta["download_url"] = "/output/" + os.path.basename(path)
    meta["uploaded"] = uploaded_name
    _set_task(tid, status="done", result=result, meta=meta,
              finished=time.time())


@app.route("/")
def index():
    return send_file(os.path.join(APP_DIR, "index.html"))


@app.route("/api/profile", methods=["GET"])
def api_profile():
    """前端拿当前档案的完整信息（向导编辑页也用它回填）。"""
    try:
        p = profile_store.load_profile()
    except ValueError as e:
        # 手写档案最常见的就是 JSON 语法错，直接把文件定位报给人看
        return jsonify({"ok": False, "error": str(e)}), 400
    if not p:
        return jsonify({"ok": True, "exists": False})
    return jsonify({"ok": True, "exists": True, "profile": p})


@app.route("/api/wizard/analyze", methods=["POST"])
def api_wizard_analyze():
    """向导第一步：上传简历 → 大模型提取档案草稿 + 生成专属模板映射。
    耗时约 30~90 秒（两次 LLM 调用），前端要给足超时并显示进度。"""
    f = request.files.get("resume")
    if f is None or not f.filename:
        return jsonify({"ok": False,
                        "error": "请先选择你的简历文件（.docx / .txt / .md）"}), 400
    low = f.filename.lower()
    if not (low.endswith(".docx") or low.endswith(".txt") or low.endswith(".md")):
        return jsonify({"ok": False,
                        "error": "只支持 .docx / .txt / .md 格式"}), 400
    from llm_client import _cfg3
    if not _cfg3("ASSIST")[0]:
        return jsonify({"ok": False,
                        "error": "还没有配置模型 key：请先点右上角「设置」填入你的 key 再解析"}), 400
    ext = "." + low.rsplit(".", 1)[-1]
    tmp = os.path.join(tempfile.gettempdir(), f"rz_wizard_{uuid.uuid4().hex}{ext}")
    f.save(tmp)
    try:
        result = wizard_mod.analyze_resume(tmp, f.filename)
        return jsonify({"ok": True, "source_file": f.filename,
                        "profile_draft": result["profile_draft"],
                        "template": result["template"]})
    except Exception as e:
        return jsonify({"ok": False,
                        "error": f"解析失败：{type(e).__name__}: {e}"}), 500
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


@app.route("/api/wizard/save", methods=["POST"])
def api_wizard_save():
    """向导第二步：保存确认后的档案并激活。"""
    data = request.get_json(silent=True) or {}
    profile = wizard_mod.normalize_profile(data)
    name = (profile.get("basic") or {}).get("name") or ""
    name = re.sub(r'[\\/:*?"<>|]', "", name).strip()
    if not name:
        return jsonify({"ok": False, "error": "姓名不能为空"}), 400
    if not profile.get("projects"):
        return jsonify({"ok": False,
                        "error": "至少要有 1 段项目经历（这是定制的原材料）"}), 400
    fname = profile_store.save_profile(f"{name}.json", profile)
    profile_store.set_active(fname)
    return jsonify({"ok": True, "file": fname,
                    "template": profile.get("template")})


@app.route("/api/wizard/template", methods=["POST"])
def api_wizard_template():
    """向导里用户手工纠正"哪几行属于哪个区域"之后重新落盘。
    版式映射以前全自动、没人核对，模型一猜错就静默不改，用户看不出来。"""
    data = request.get_json(silent=True) or {}
    pack = str(data.get("pack") or "")
    regions = data.get("regions") or {}
    if not isinstance(regions, dict):
        return jsonify({"ok": False, "error": "regions 格式不对"}), 400
    try:
        out = wizard_mod.update_template(pack, regions)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False,
                        "error": f"保存映射失败：{type(e).__name__}: {e}"}), 500
    return jsonify(dict(out, ok=True))


@app.route("/api/settings/llm", methods=["GET"])
def api_settings_get():
    """设置页读取：返回当前生效的模型配置（key 只给尾 4 位）。"""
    from llm_client import _cfg3
    def masked(prefix):
        k, b, m = _cfg3(prefix)
        return {"base": b or "", "model": m or "",
                "configured": bool(k),
                "key_tail": (k[-4:] if k and len(k) >= 8 else "")}
    return jsonify({"ok": True, "primary": masked("ASSIST"), "backup": masked("BACKUP")})


@app.route("/api/settings/llm", methods=["POST"])
def api_settings_save():
    """设置页保存：key 留空 = 保持原值不变。写入 profiles/api_key.json。"""
    data = request.get_json(silent=True) or {}
    path = os.path.join(APP_DIR, "profiles", "api_key.json")
    cur = {}
    try:
        with open(path, encoding="utf-8") as f:
            cur = json.load(f)
        if not isinstance(cur, dict):
            cur = {}
    except Exception:
        cur = {}
    for prefix, block_key in (("ASSIST", "primary"), ("BACKUP", "backup")):
        block = data.get(block_key) or {}
        if not isinstance(block, dict):
            continue
        for field, suffix in (("api_key", "API_KEY"), ("base", "BASE_URL"), ("model", "MODEL")):
            v = str(block.get(field) or "").strip()
            if v:
                cur[prefix + "_" + suffix] = v
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cur, f, ensure_ascii=False, indent=2)
    return jsonify({"ok": True})


USER_TPL_RE = re.compile(r"^user_[0-9a-f]{8}$")   # 向导生成的专属模板包名


def _active_template_name():
    try:
        return (profile_store.load_profile() or {}).get("template")
    except Exception:
        return None


@app.route("/api/templates", methods=["GET"])
def api_templates_list():
    """列出向导为上传的简历生成的专属模板包（包内含用户简历原件，占地方也涉隐私）。"""
    out = []
    for name in template_export.list_templates():
        if not USER_TPL_RE.match(name):
            continue
        d = os.path.join(template_export.TEMPLATES_DIR, name)
        try:
            with open(os.path.join(d, "template.json"), encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
        out.append({"name": name,
                    "display_name": cfg.get("display_name") or name,
                    "date": time.strftime("%Y-%m-%d", time.localtime(os.path.getmtime(d)))})
    return jsonify({"ok": True, "templates": out, "in_use": _active_template_name()})


@app.route("/api/templates/<name>", methods=["DELETE"])
def api_templates_delete(name):
    """删掉一个专属模板包。正则已挡住路径穿越；当前档案在用的不允许删。"""
    if not USER_TPL_RE.match(name):
        return jsonify({"ok": False, "error": "只能删除向导生成的专属模板（内置模板不能删）"}), 400
    if _active_template_name() == name:
        return jsonify({"ok": False,
                        "error": "这个模板正被当前档案使用，先重新解析简历换掉它再删"}), 400
    d = os.path.join(template_export.TEMPLATES_DIR, name)
    if not os.path.isdir(d):
        return jsonify({"ok": False, "error": "该模板包不存在，可能已经删过了"}), 404
    shutil.rmtree(d)
    return jsonify({"ok": True})


@app.route("/api/ping", methods=["GET", "POST"])
def api_ping():
    """诊断端点：从服务器进程内直接调模型，排查 API 层问题。

    可带 POST body（设置页当前表单值）：这样"填完 key 先测再存"测的是
    刚填的那个 key，而不是磁盘上还没被覆盖的旧配置。body 里留空的字段
    回落到已保存的配置（设置页的 key 输入框本来就是"留空 = 不修改"）。
    """
    from llm_client import _cfg3, _call_once
    body = request.get_json(silent=True) or {}
    t0 = time.time()

    def probe(prefix, block_key):
        k, b, m = _cfg3(prefix)
        blk = body.get(block_key) or {}
        if isinstance(blk, dict):
            k = str(blk.get("api_key") or "").strip() or k
            b = str(blk.get("base") or "").strip() or b
            m = str(blk.get("model") or "").strip() or m
        out = {"base": b or "", "model": m or "", "key_len": len(k) if k else 0,
               "configured": bool(k and b and m), "ok": False, "status": ""}
        if not out["configured"]:
            out["status"] = "未配置完整（key / 地址 / 模型名都要填）"
            return out
        # ⚠️ max_tokens 不能太小：主模型是推理模型，思考 token 会先把预算吃光，
        # 给 30 会导致 content 为空被误判成"模型坏了"。实测给 400 才稳。
        content, usage = _call_once(k, b, m,
                                    [{"role": "user", "content": "回复两个字：收到"}],
                                    400, temperature=0.6)
        if content:
            out["ok"] = True
            out["status"] = "连通正常"
        elif isinstance(usage, dict) and usage.get("_rl"):
            out["status"] = "被限流（429），等一两分钟再试"
        else:
            out["status"] = "调不通：检查 key、地址、模型名是否正确"
        return out

    prim = probe("ASSIST", "primary")
    back = probe("BACKUP", "backup")
    return jsonify({"ok": True,
                    "elapsed": round(time.time() - t0, 1),
                    "primary": prim, "backup": back})


def _read_text_any_codec(path):
    """读 txt/md：先按 utf-8，失败回退 gbk（中文 Windows 记事本默认 ANSI/GBK），
    都不行再带替换符硬读——绝不让乱码素材悄悄喂给模型。"""
    raw = open(path, "rb").read()
    for codec in ("utf-8", "gbk"):
        try:
            return raw.decode(codec)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _format_parsed(parsed):
    """
    把上传简历的解析结果格式化成喂给 LLM 的文本。
    ⚠️ 不含 fields（姓名/电话/邮箱/籍贯/学历等）——这些由用户档案（profiles/）提供，
    解析出来再喂一遍是重复浪费 token，且容易和档案冲突。
    只喂项目、技能、教育、意向这类"可能有新信息"的内容。
    """
    parts = []
    secs = parsed.get("sections") or {}
    LABEL = {"education": "教育背景", "objective": "求职意向",
             "projects": "项目经历", "skills": "专业技能"}
    for key, label in LABEL.items():
        txt = (secs.get(key) or "").strip()
        if txt:
            parts.append(f"\n【{label}】\n{txt}")
    return "\n".join(parts)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(silent=True) or {}
    jd = (data.get("jd") or request.form.get("jd") or "").strip()
    focus = (data.get("focus") or request.form.get("focus") or "").strip()

    if len(jd) < 20:
        return jsonify({"ok": False,
                        "error": "岗位描述太短了，至少把任职要求那几行贴进来"}), 400

    # 没有 key 时给可操作的引导，而不是让任务跑半天才报"模型不可用"
    from llm_client import _cfg3
    if not _cfg3("ASSIST")[0]:
        return jsonify({"ok": False,
                        "error": "还没有配置模型 key：请点右上角「设置」填入你的 key"
                                 "（商汤 sensenova 官网注册即可获得）"}), 400

    # --- 解析上传的简历文件（可选）---
    extra_material = ""
    uploaded_name = None
    f = request.files.get("resume")
    if f is not None and f.filename:
        # ⚠️ 不能用 secure_filename 做临时文件名：它会把纯中文文件名剥得只剩扩展名
        #（"张三简历.docx" → "docx"），并发上传会撞同一个临时路径。用 uuid 保证唯一。
        # 扩展名白名单直接对原始文件名判断。
        low = f.filename.lower()
        if not (low.endswith(".docx") or low.endswith(".txt")
                or low.endswith(".md")):
            return jsonify({"ok": False,
                            "error": "只支持 .docx / .txt / .md 格式的简历文件"}), 400
        ext = "." + low.rsplit(".", 1)[-1]
        tmp = os.path.join(tempfile.gettempdir(), f"rz_upload_{uuid.uuid4().hex}{ext}")
        f.save(tmp)
        try:
            if low.endswith(".docx"):
                parsed = parse_resume_docx(tmp)
                extra_material = _format_parsed(parsed)
            else:
                extra_material = _read_text_any_codec(tmp)
            uploaded_name = f.filename
        except Exception as e:
            return jsonify({"ok": False,
                            "error": f"简历文件解析失败：{type(e).__name__}: {e}"}), 400
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    # 无档案时给出可操作的错误（引导完成首次配置，M2 会做向导页）
    try:
        profile = profile_store.load_profile()
    except ValueError as e:
        return jsonify({"ok": False, "error": f"个人档案读不了：{e}"}), 400
    if not profile:
        return jsonify({"ok": False,
                        "error": "还没有个人档案。请先把你的资料 JSON 放进 profiles/ 目录"
                                 "（参考 profiles/示例-张三.json 的结构）"}), 400

    # 生成 + 导出放后台线程，这里立刻返回 task_id
    tid = _new_task()
    threading.Thread(
        target=_run_task,
        args=(tid, jd, focus, extra_material, uploaded_name, profile),
        daemon=True,
    ).start()

    return jsonify({"ok": True, "task_id": tid})


@app.route("/api/status/<tid>", methods=["GET"])
def api_status(tid):
    """前端轮询这个接口查生成进度。"""
    t = _get_task(tid)
    if not t:
        return jsonify({"ok": False, "status": "gone",
                        "error": "任务不存在或已失效（服务是不是重启过？）"}), 404

    st = t.get("status")
    if st == "done":
        return jsonify({"ok": True, "status": "done",
                        "result": t["result"], "meta": t["meta"]})
    if st == "error":
        return jsonify({"ok": False, "status": "error",
                        "error": t.get("error") or "生成失败",
                        "meta": t.get("meta")})

    waited = time.time() - t.get("started", t.get("created", time.time()))
    return jsonify({"ok": True, "status": st, "waited": round(waited, 1)})


@app.errorhandler(404)
def _404(_):
    return jsonify({"ok": False, "error": "找不到该地址"}), 404


@app.errorhandler(500)
def _500(e):
    return jsonify({"ok": False, "error": "服务器内部错误"}), 500


if __name__ == "__main__":
    print("=" * 56)
    print("  简历生成器  启动中")
    print(f"  打开浏览器访问： http://127.0.0.1:5050")
    print(f"  生成文件目录：   {OUTPUT_DIR}")
    print("=" * 56)
    app.run(host="127.0.0.1", port=5050, debug=False, threaded=True)
