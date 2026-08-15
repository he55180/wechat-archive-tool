#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI增强分类模块（可选增值功能）
-----------------------------------------------
设计原则：
1. 只在规则分类"拿不准"时触发（低置信度 or 完全未命中），大多数文章 AI 不介入。
2. 默认关闭：config.json ai_classify.enabled=false 或 api_key 为空时直接走规则降级。
3. 熔断器：单次批处理内连续失败 >= circuit_breaker_threshold 后自动跳过，
   避免 API 宕机时每篇都等超时。
4. 任何失败最终都回退到规则引擎次优解，不会比纯规则版本更差。

使用方式（在 app.py 调用）：
    from ai_classify import CircuitBreaker, classify_with_ai, RuleResult

    circuit_breaker = CircuitBreaker(threshold=config["ai_classify"].get("circuit_breaker_threshold", 5))
    result = classify_with_ai(title, content, rule_result, config, circuit_breaker, VALID_CATEGORIES)
"""

import logging
import requests

log = logging.getLogger(__name__)

# ── 提示词模板 ─────────────────────────────────────────────────────
_CLASSIFY_SYSTEM = (
    "你是一个文档分类助手。根据文章标题和摘要，从以下分类列表中选出最合适的一个分类，"
    "只输出分类名称，不要任何解释或多余文字。"
)


# ──────────────────────────────────────────────────────────────────
# 数据结构
# ──────────────────────────────────────────────────────────────────

class RuleResult:
    """
    封装规则引擎的分类结果。

    Attributes
    ----------
    category : str
        规则命中的分类文件夹名（若完全未命中则为 best_guess）
    best_guess : str
        规则次优解（兜底分类文件夹名，通常等于 DEFAULT_FOLDER）
    tags : list[str]
        对应的标签列表
    is_confident : bool
        规则引擎是否"有把握"：命中关键词 >= 1 个时视为有把握
    hit_count : int
        命中关键词数量（用于置信度判断）
    """

    def __init__(self, category: str, best_guess: str, tags: list, hit_count: int, matched_keywords: list = None):
        self.category         = category
        self.best_guess       = best_guess
        self.tags             = tags
        self.hit_count        = hit_count
        self.matched_keywords = matched_keywords if matched_keywords is not None else []
        # 命中 >= 1 个关键词即视为"有把握"，只有完全0命中才触发AI二次判断
        self.is_confident = hit_count >= 1


# ──────────────────────────────────────────────────────────────────
# 熔断器
# ──────────────────────────────────────────────────────────────────

class CircuitBreaker:
    """
    简单计数熔断器，只在本次程序运行期间有效，不持久化。

    连续失败次数达到 threshold 后，is_open() 返回 True，
    此后 classify_with_ai() 将直接走规则降级，不再等待超时。
    """

    def __init__(self, threshold: int = 5):
        self._threshold       = threshold
        self._consecutive_fail = 0
        self._opened          = False

    def is_open(self) -> bool:
        return self._opened

    def record_success(self):
        self._consecutive_fail = 0
        # 不自动关闭熔断器（单次运行内保守策略，有一次成功也不重置）

    def record_failure(self):
        self._consecutive_fail += 1
        if self._consecutive_fail >= self._threshold:
            if not self._opened:
                log.warning(
                    "AI分类熔断器触发：连续失败 %d 次，本次运行剩余文章将直接使用规则分类。",
                    self._consecutive_fail,
                )
            self._opened = True


# ──────────────────────────────────────────────────────────────────
# 核心 AI 调用
# ──────────────────────────────────────────────────────────────────

def _call_ai_classify(
    title: str,
    content_snippet: str,
    valid_categories: list[str],
    config: dict,
    timeout: float,
) -> str:
    """
    向 AI 发起分类请求，返回分类名称字符串。
    失败时抛异常（由调用方捕获）。
    """
    ai_cfg   = config.get("ai_classify", {})
    api_key  = ai_cfg.get("api_key", "").strip()
    api_base = ai_cfg.get("api_base", "https://api.deepseek.com/v1").rstrip("/")
    model    = ai_cfg.get("model", "deepseek-chat")

    cat_list = "\n".join(f"- {c}" for c in valid_categories)
    user_msg = (
        f"文章标题：{title}\n\n"
        f"文章摘要（前500字）：{content_snippet[:500]}\n\n"
        f"可选分类：\n{cat_list}"
    )

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _CLASSIFY_SYSTEM},
            {"role": "user",   "content": user_msg},
        ],
        "max_tokens": 30,
        "temperature": 0.0,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    resp = requests.post(
        f"{api_base}/chat/completions",
        json=payload,
        headers=headers,
        timeout=timeout,
    )

    if resp.status_code != 200:
        raise RuntimeError(f"API HTTP {resp.status_code}: {resp.text[:200]}")

    data = resp.json()
    result = data["choices"][0]["message"]["content"].strip()
    if not result:
        raise ValueError("AI返回内容为空")
    return result


# ──────────────────────────────────────────────────────────────────
# 对外主接口
# ──────────────────────────────────────────────────────────────────

def classify_with_ai(
    title: str,
    content: str,
    rule_result: RuleResult,
    config: dict,
    circuit_breaker: CircuitBreaker,
    valid_categories: list[str],
) -> tuple[str, list]:
    """
    对单篇文章执行 AI 增强分类。

    Parameters
    ----------
    title : str
        文章标题
    content : str
        文章正文（Markdown 格式）
    rule_result : RuleResult
        规则引擎输出的结果对象
    config : dict
        完整 config.json 解析后的字典
    circuit_breaker : CircuitBreaker
        熔断器实例（应在批处理任务开始时创建，整批共用一个）
    valid_categories : list[str]
        合法分类文件夹名列表（AI 返回不在此列表内的值视为无效）

    Returns
    -------
    (folder_name: str, tags: list[str])
        最终确定的分类文件夹名和标签列表
    """
    ai_cfg = config.get("ai_classify", {})

    # ── 1. 规则有把握 → 直接用规则结果，AI 完全不介入 ────────────
    if rule_result.is_confident:
        return rule_result.category, rule_result.tags

    # ── 2. 快速降级路径 ──────────────────────────────────────────
    if not ai_cfg.get("enabled", False):
        return rule_result.best_guess, rule_result.tags

    api_key = ai_cfg.get("api_key", "").strip()
    if not api_key:
        return rule_result.best_guess, rule_result.tags

    if circuit_breaker.is_open():
        log.warning("AI分类已熔断，本篇直接使用规则次优解: %s", rule_result.best_guess)
        return rule_result.best_guess, rule_result.tags

    # ── 3. 调用 AI，带主超时 + 一次重试（更短超时） ────────────────
    primary_timeout = float(ai_cfg.get("timeout_seconds", 5))
    retry_timeout   = min(primary_timeout * 0.6, 3.0)   # 重试用更短超时
    max_retries     = int(ai_cfg.get("max_retries", 1))

    last_error = None
    for attempt in range(1 + max_retries):
        timeout = primary_timeout if attempt == 0 else retry_timeout
        try:
            ai_category = _call_ai_classify(
                title, content, valid_categories, config, timeout
            )
            # ── 校验返回是否在合法列表里 ─────────────────────────
            if ai_category not in valid_categories:
                raise ValueError(f"AI返回分类不在预设列表内: '{ai_category}'")

            circuit_breaker.record_success()
            log.info("AI分类成功（第%d次尝试）: %s → %s", attempt + 1, title[:30], ai_category)

            # 推断对应标签（取文件夹名中的中文部分）
            tag = ai_category.split(".")[-1] if "." in ai_category else ai_category
            return ai_category, [tag]

        except requests.exceptions.Timeout:
            last_error = f"超时（{timeout}s）"
        except requests.exceptions.RequestException as e:
            last_error = f"网络异常: {e}"
        except Exception as e:
            last_error = str(e)
            # 非网络类错误不重试（如返回格式问题）
            break

    # ── 4. 所有尝试均失败 → 熔断计数 + 回退规则次优解 ──────────────
    circuit_breaker.record_failure()
    log.warning("AI二次分类失败，回退规则次优解：%s | 原因：%s", rule_result.best_guess, last_error)
    return rule_result.best_guess, rule_result.tags
