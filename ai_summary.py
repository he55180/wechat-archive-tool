#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI摘要模块（可选增值功能）
-----------------------------------------------
默认关闭：config.json 中 ai_summary.enabled=false 或 api_key 为空时，
generate_summary() 始终返回 None，不发任何网络请求，不影响主流程。

异常策略：
- 任何网络错误、超时、API 非200、返回内容为空 → 返回 None + log.warning
- 绝不向上抛异常，不能中断抓取→排版→归档主流程。
"""

import logging
import requests

log = logging.getLogger(__name__)

# ── 系统提示词（对 OpenAI 兼容接口） ──────────────────────────────
_SYSTEM_PROMPT = (
    "你是一个专业的内容摘要助手。"
    "请用一句话（不超过60个中文字）概括以下文章的核心内容，"
    "语言精练，不要带任何标点以外的多余文字，不要加前缀说明。"
)


def generate_summary(content: str, config: dict) -> str | None:
    """
    生成文章一句话摘要。

    Parameters
    ----------
    content : str
        文章正文（Markdown 格式）
    config : dict
        完整 config.json 解析后的字典

    Returns
    -------
    str | None
        成功返回摘要字符串；
        未启用 / 未配置 / 任何失败 → 返回 None，不抛异常。
    """
    ai_cfg = config.get("ai_summary", {})

    # ── 快速跳过检查 ─────────────────────────────────────────────
    if not ai_cfg.get("enabled", False):
        return None

    api_key = ai_cfg.get("api_key", "").strip()
    if not api_key:
        return None

    # ── 读取参数，设置合理默认值 ──────────────────────────────────
    api_base      = ai_cfg.get("api_base", "https://api.deepseek.com/v1").rstrip("/")
    model         = ai_cfg.get("model", "deepseek-chat")
    timeout_sec   = int(ai_cfg.get("timeout_seconds", 8))
    max_chars     = int(ai_cfg.get("max_content_chars", 3000))

    # 正文截断，控制 token 成本
    truncated = content[:max_chars] if len(content) > max_chars else content

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user",   "content": truncated},
        ],
        "max_tokens": 120,
        "temperature": 0.3,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(
            f"{api_base}/chat/completions",
            json=payload,
            headers=headers,
            timeout=timeout_sec,
        )
    except requests.exceptions.Timeout:
        log.warning("AI摘要生成失败，已跳过：请求超时（超过 %ds）", timeout_sec)
        return None
    except requests.exceptions.RequestException as e:
        log.warning("AI摘要生成失败，已跳过：网络异常 - %s", e)
        return None

    # ── 处理 HTTP 层错误 ──────────────────────────────────────────
    if resp.status_code != 200:
        log.warning(
            "AI摘要生成失败，已跳过：API 返回 HTTP %d - %s",
            resp.status_code, resp.text[:200],
        )
        return None

    # ── 解析返回内容 ──────────────────────────────────────────────
    try:
        data = resp.json()
        summary = data["choices"][0]["message"]["content"].strip()
        if not summary:
            raise ValueError("摘要内容为空")
        return summary
    except Exception as e:
        log.warning("AI摘要生成失败，已跳过：解析响应失败 - %s", e)
        return None
