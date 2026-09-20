# -*- coding: utf-8 -*-
"""
LLM 配置模板：把本文件复制为 llm_config.py 并填入你自己的 key（不要提交到 git）。
更推荐的方式：直接在网页右上角「设置」里填，key 会存在 profiles/api_key.json，
本文件可以完全不建。

主模型（必填）"""
ASSIST_API_KEY = ""                                  # 你的 API key
ASSIST_BASE_URL = "https://token.sensenova.cn/v1"    # 商汤开放平台地址
ASSIST_MODEL = "sensenova-6.8-flash-lite"

# 备用模型（选填：主模型限流时自动顶上）
BACKUP_API_KEY = ""
BACKUP_BASE_URL = ""
BACKUP_MODEL = ""

"""
去哪申请：
  商汤 sensenova：https://platform.sensenova.cn 注册后在控制台创建 API key
  （任何 OpenAI 兼容接口的模型服务都可以，改 BASE_URL 和 MODEL 即可）
"""
