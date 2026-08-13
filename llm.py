"""DeepSeek 客户端：批量段落分类。

设计目标：把若干段落文本批量送入大模型，让模型为每段输出一个
样式标签（H1~H9 / BODY / BULLET / NUM / TABLE_HEADER / TABLE_BODY /
CODE / IMAGE / CAPTION），下游 typeset.py 据此赋样式。
"""
from __future__ import annotations

import json
import re
from typing import List

import requests

from config import Config

# 允许的标签集合（与 config.yaml style_map 的 key 保持一致）
ALLOWED_LABELS = {
    "H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9",
    "BODY", "BULLET", "NUM",
    "TABLE_HEADER", "TABLE_BODY", "CODE", "IMAGE", "CAPTION",
}

SYSTEM_PROMPT = (
    "你是一个 Word 文档排版分类助手。给定若干段落文本，"
    "为每一段输出一个样式标签，只能从以下集合中选择：\n"
    "H1~H9: 各级标题（按层级）；\n"
    "BODY: 正文段落；\n"
    "BULLET: 无序列表项（圆点）；\n"
    "NUM: 有序列表项（数字编号）；\n"
    "TABLE_HEADER: 表格表头行；\n"
    "TABLE_BODY: 表格正文行；\n"
    "CODE: 代码块/等宽字体段；\n"
    "IMAGE: 图片所在段；\n"
    "CAPTION: 图表题注。\n"
    "只允许返回上述标签，禁止返回其他文字。"
)


def _build_user_prompt(items: List[str], fallback_label: str) -> str:
    lines = [
        "请对下面每一段输出一个标签，严格输出 JSON："
        '{"labels":["H1","BODY",...]}',
        "labels 数组长度必须与下方段落数一致，顺序对应。",
    ]
    for i, t in enumerate(items, 1):
        lines.append(f"{i}. {t}")
    lines.append(f'若无法判断，统一返回 "{fallback_label}"。')
    return "\n".join(lines)


def _parse_labels(text: str, n: int, fallback: str) -> List[str]:
    """从模型输出中解析 n 个标签，缺失或非法项用 fallback 补齐。"""
    # 1) 优先从 JSON 中提取
    try:
        m = re.search(r"\{.*\}", text, re.S)
        if m:
            obj = json.loads(m.group(0))
            labels = obj.get("labels", [])
            if isinstance(labels, list) and len(labels) == n:
                return [str(x).upper() if str(x).upper() in ALLOWED_LABELS else fallback
                        for x in labels]
    except Exception:
        pass
    # 2) 兜底：按行解析
    out: List[str] = []
    for line in text.splitlines():
        for tok in re.findall(r"[A-Za-z0-9_]+", line):
            tok = tok.upper()
            if tok in ALLOWED_LABELS:
                out.append(tok)
                break
    while len(out) < n:
        out.append(fallback)
    return out[:n]


def classify_paragraphs(cfg: Config, items: List[str]) -> List[str]:
    """批量分类，返回与 items 等长的标签列表。"""
    if not items:
        return []
    fallback = cfg.typeset.get("fallback_label", "BODY")
    batch = cfg.llm.get("batch_size", 20)
    max_len = cfg.typeset.get("max_text_len", 500)
    api_key = cfg.api_key
    results: List[str] = []

    for i in range(0, len(items), batch):
        chunk = [t[:max_len] for t in items[i:i + batch]]
        payload = {
            "model": cfg.llm["model"],
            "temperature": cfg.llm.get("temperature", 0.0),
            "max_tokens": cfg.llm.get("max_tokens", 1024),
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": _build_user_prompt(chunk, fallback)},
            ],
        }
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        resp = requests.post(
            f"{cfg.llm['base_url']}/chat/completions",
            json=payload,
            headers=headers,
            timeout=cfg.llm.get("timeout", 30),
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
        results.extend(_parse_labels(content, len(chunk), fallback))
    return results
