# -*- coding: utf-8 -*-
"""
LLM 调用模块
  - 主 provider：商汤 token.sensenova.cn / sensenova-6.8-flash-lite（实测 45~145s）
  - 备 provider：小米 mm-api / mimo-v2.5（⚠️ 实测经常 503 model_not_found，不可靠）
  - 每次调用都先试主模型，主模型失败才走备胎
    （2026-09-10 去掉了"429 冷却期跳过主模型"的逻辑，原因见 call_model 的注释）
  - 两家都挂 → 返回 None，由调用方降级

配置读取优先级（2026-09-16 更新）：
  0. profiles/api_key.json  网页「设置」页写入（每台电脑自己的，最高优先级）
  1. llm_config.py          文件方式配置（参考 llm_config.example.py）
  2. 环境变量               最终兜底
"""

import json
import os
import sys
import time
import requests

_HERE = os.path.dirname(os.path.abspath(__file__))

# 每台电脑自己的 key 设置（网页「设置」页写入，2026-09-16 新增）。
# 优先级最高：这样发包时 llm_config.py 可以不带任何 key，收包人填这里即可。
_KEY_SETTINGS = os.path.join(_HERE, "profiles", "api_key.json")


def _read_key_settings():
    try:
        with open(_KEY_SETTINGS, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}

TIMEOUT = (10, 900)       # 连接 10 秒 / 读取 900 秒
                          # ⚠️ 踩过的坑（2026-09-10 晚）：
                          #   曾把读超时改成 300 秒，结果实测出现「思考超过 300 秒」的正常生成，
                          #   被超时掐断 → 用户看到"模型不可用"。
                          #   模板版要求输出 6 条项目要点 + 6 条技能，内容量比原来大，
                          #   模型思考时间也跟着变长。所以这里必须留足余量。
                          #   宁可让用户多等，也不要掐断本来能成功的生成。


def _add_path(p):
    if p and p not in sys.path:
        sys.path.insert(0, p)


def _cfg3(prefix):
    """读 <prefix>_API_KEY / _BASE_URL / _MODEL。优先级：
    1. profiles/api_key.json（设置页写入，每台电脑自己的）
    2. llm_config.py（参考 llm_config.example.py 的文件方式）
    3. 环境变量（最终兜底）
    """
    _add_path(_HERE)
    # 0) 设置页写入的 key（最高优先级，逐字段合并）
    s = _read_key_settings()
    key = s.get(prefix + "_API_KEY")
    base = s.get(prefix + "_BASE_URL")
    model = s.get(prefix + "_MODEL")
    # 1) 本目录专属配置
    try:
        import llm_config  # noqa
        key = key or getattr(llm_config, prefix + "_API_KEY", None)
        base = base or getattr(llm_config, prefix + "_BASE_URL", None)
        model = model or getattr(llm_config, prefix + "_MODEL", None)
    except ImportError:
        pass
    # 2) 环境变量最终兜底
    key = key or os.environ.get(prefix + "_API_KEY")
    base = base or os.environ.get(prefix + "_BASE_URL")
    model = model or os.environ.get(prefix + "_MODEL")
    return key, base, model


def _call_once(key, base, model, messages, max_tokens, temperature=0.6,
               quiet=False, no_think=False):
    """
    单次调用。返回 (content, usage)；429 → (None, {'_rl': 1})；失败 → (None, None)。
    quiet=False 时把失败原因 print 出来（排查用），不影响返回值。

    no_think=True 时传 thinking={"type":"disabled"}，关掉推理模型的"思考"。
      ⚠️ 实测（2026-09-10，sensenova-6.8-flash-lite）：
         不传        → 48.4 秒，思考 1345 token
         no_think    → 13.5 秒，思考 0 token       ← 快 3.6 倍
         传 reasoning_effort="low"  → 无效（47.1 秒）
         传 enable_thinking=False   → 被 429 拒
      慢的根源就是"思考"：这个模型会用满预算去想，
      所以关掉思考是最有效的提速手段。
    """
    payload = {
        "model": model,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "stream": False,
    }
    if no_think:
        payload["thinking"] = {"type": "disabled"}

    try:
        resp = requests.post(
            base.rstrip("/") + "/chat/completions",
            headers={"Authorization": "Bearer " + key},
            json=payload,
            timeout=TIMEOUT,
        )
        if resp.status_code == 429:
            return None, {"_rl": 1}
        if resp.status_code != 200:
            if not quiet:
                print(f"[llm] HTTP {resp.status_code} from {base} "
                      f"model={model}: {resp.text[:180]}", flush=True)
            return None, None
        d = resp.json()
        content = (d.get("choices", [{}])[0].get("message", {})
                   .get("content") or "").strip()
        if not content:  # 推理模型预算烧尽返回空串，视为失败交给备胎
            return None, None
        return content, (d.get("usage") or {})
    except Exception as e:
        if not quiet:
            print(f"[llm] 异常 {type(e).__name__}: {e}  base={base} model={model}",
                  flush=True)
        return None, None


def call_model(messages, max_tokens=3000, temperature=0.6, no_think=True):
    """
    返回 (content, usage, model_used)。全挂返回 (None, None, None)。

    2026-09-10 改动一：去掉"主模型撞 429 后 60 秒冷却、冷却期内跳过主模型"的逻辑。
    原因：429 响应很快（<1 秒），跳过主模型几乎省不下时间；
    而备胎 mimo-v2.5 实测经常 503 不可用，一旦进入冷却期，
    这 60 秒内所有请求都必然失败（实测 1.4 秒就返回"模型全部不可用"），
    用户会以为工具坏了。现在每次都先试主模型 —— 撞 429 也只损失不到 1 秒。

    2026-09-10 改动二：主模型默认 no_think=True（关掉推理模型的"思考"）。
    实测同一 prompt：不关思考 48.4 秒、关掉 13.5 秒（快 3.6 倍）。
    这是把生成时间从 17 分钟压下来的关键手段。
    ⚠️ 备胎不传这个参数（不同 provider 的参数约定不一样，传了可能 400）。
    """
    pkey, pbase, pmodel = _cfg3("ASSIST")
    if pkey and pbase and pmodel:
        content, usage = _call_once(pkey, pbase, pmodel,
                                    messages, max_tokens, temperature,
                                    no_think=no_think)
        if content is not None:
            return content, usage, pmodel
    bkey, bbase, bmodel = _cfg3("BACKUP")
    if bkey and bbase and bmodel:
        content, usage = _call_once(bkey, bbase, bmodel,
                                    messages, max_tokens, temperature)  # 备胎不关思考
        if content is not None:
            return content, usage, bmodel
    return None, None, None


if __name__ == "__main__":
    # 快速连通性自检
    print("主 provider:", _cfg3("ASSIST")[1], _cfg3("ASSIST")[2])
    print("备 provider:", _cfg3("BACKUP")[1], _cfg3("BACKUP")[2])
    c, u, m = call_model(
        [{"role": "user", "content": "回复两个字：收到"}], max_tokens=20)
    print("model:", m, "| content:", c, "| usage:", u)
