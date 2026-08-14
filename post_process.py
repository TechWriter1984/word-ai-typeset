"""步骤4 + 增强后处理：合并完成后的收尾工作。

包含：
1. remove_all_section_markers  —— 删除每个节的 [SectionName: XXX] 标记行（问题4上半部分）
2. collect_bodytext_block_elements —— 返回 BodyText 节范围内的顶级块（w:p / w:tbl），
   供 format_images 限定处理范围，避免封面 Logo / Symbol 表 Icon 被误处理（问题1）
3. blacken_all_table_text —— 所有表格文字颜色改为黑色（问题4下半部分 + 问题5 中留黑的需求）
4. replace_placeholders_with_llm —— 大模型分析排版后正文，替换模板遗留的 {{XXX}} 占位符（问题4）
5. update_terms_table —— 移植 VBA 更新术语表，双语种搜索，删未用行+保留行变黑（问题5）
6. run_post_process —— 一键按开关顺序执行全部子步骤
"""
from __future__ import annotations

import re
from typing import List, Tuple

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table as DocxTable

from config import Config
# llm 函数在本文件内按需使用 requests 直接调用（见 replace_placeholders）


# ========== 通用辅助 ==========

def _para_text(p_el) -> str:
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def _set_run_color_black(p_el):
    """段落内所有 run 字体颜色设黑。"""
    for r in p_el.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = r.makeelement(qn("w:rPr"), {})
            r.insert(0, rPr)
        color = rPr.find(qn("w:color"))
        if color is None:
            color = rPr.makeelement(qn("w:color"), {})
            rPr.append(color)
        color.set(qn("w:val"), "000000")


# ========== 1. 删除所有 [SectionName: XXX] 标记行 ==========

SECTION_MARKER_RE = re.compile(r"\[SectionName\s*:\s*[^\]]+\]")


def remove_all_section_markers(doc: Document) -> dict:
    """删除正文 body 中所有包含 [SectionName: XXX] 的段落（整段删除）。

    返回: {"removed": n}
    """
    body = doc.element.body
    removed = 0
    for p in list(body.findall(qn("w:p"))):
        text = _para_text(p)
        if SECTION_MARKER_RE.search(text):
            body.remove(p)
            removed += 1
    return {"removed": removed}


# ========== 2. 收集 BodyText 节范围内的顶级块 ==========

def collect_bodytext_block_elements(doc: Document, section_marker: str) -> list:
    """返回模板中 BodyText 节对应的顶级块（w:p / w:tbl）列表（不含 sectPr）。

    节范围界定：
      起点：[SectionName: {section_marker}] 标记段落的下一个兄弟
      终点：遇到 [SectionName: ...] 类标记或 body 末尾或 sectPr
    若找不到 marker，返回空列表。
    """
    body = doc.element.body
    marker_tag = f"[SectionName: {section_marker}]"
    marker_p = None
    for p in body.findall(qn("w:p")):
        txt = _para_text(p)
        if marker_tag in txt:
            marker_p = p
            break
    if marker_p is None:
        return []

    blocks: list = []
    started = False
    for child in list(body):
        if child is marker_p:
            started = True
            continue
        if not started:
            continue
        if child.tag == qn("w:sectPr"):
            break
        if child.tag == qn("w:p"):
            txt = _para_text(child)
            # 遇到下一个 SectionName 标记即停止（BodyText 后若还有其他节）
            if SECTION_MARKER_RE.search(txt):
                break
            blocks.append(child)
        elif child.tag == qn("w:tbl"):
            blocks.append(child)
    return blocks


# ========== 3. 所有表格文字颜色统一黑色 ==========

def blacken_all_table_text(doc: Document) -> dict:
    """对 doc.tables 中每个单元格的每个段落内所有 run 字体颜色设黑。"""
    tables_done = 0
    cells_done = 0
    for tbl in doc.tables:
        tables_done += 1
        for row in tbl.rows:
            for cell in row.cells:
                cells_done += 1
                for p in cell.paragraphs:
                    _set_run_color_black(p._p)
    return {"tables": tables_done, "cells": cells_done}


# ========== 4. {{XXX}} 占位符 LLM 替换 ==========

PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_\-\u4e00-\u9fa5]*)\s*\}\}")


def collect_placeholders(doc: Document) -> List[Tuple[str, set]]:
    """遍历 doc 全部段落+表格段落，收集未重复占位符集合，并返回 (正文文本摘要, {placeholder,...})。"""
    placeholders: set = set()
    body_text_parts: List[str] = []

    def scan_p_el(p_el, for_placeholder=True):
        txt = _para_text(p_el)
        if for_placeholder:
            for m in PLACEHOLDER_RE.finditer(txt):
                placeholders.add(m.group(1))
        return txt

    # 顶层段落
    for p in doc.element.body.findall(qn("w:p")):
        body_text_parts.append(scan_p_el(p))
    # 表格段落
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    scan_p_el(p._p)
    body_snippet = "\n".join(
        (t.strip() for t in body_text_parts if t and t.strip())
    )[:4000]
    return body_snippet, placeholders


def _llm_infer_placeholder_values(cfg: Config, body_text: str, placeholders: List[str]) -> dict:
    """调用大模型，基于正文内容推断 {{XXX}} 占位符的真实值。返回 {placeholder: value}。"""
    if not placeholders:
        return {}
    llm_cfg = cfg.llm
    api_key = cfg.api_key

    system = (
        "你是一个 Word 文档占位符填充助手。用户会给你一段排版完成的文档正文文本片段，"
        "以及一组占位符名称。请阅读正文内容，推断每个占位符应该填入的实际数据。"
        "严格以 JSON 格式返回：{\"占位符名1\": \"值1\", \"占位符名2\": \"值2\"}。"
        "如果某个占位符从正文无法推断，填空字符串 \"\"，不要胡编。"
    )
    user_lines = ["正文文本片段："]
    user_lines.append(body_text)
    user_lines.append("\n需要填充的占位符列表：" + ", ".join(placeholders))
    user_lines.append("\n请仅返回 JSON，不要有其他文字。")
    user_prompt = "\n".join(user_lines)

    payload = {
        "model": llm_cfg["model"],
        "temperature": 0.0,
        "max_tokens": llm_cfg.get("max_tokens", 1024),
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user_prompt},
        ],
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    import requests
    try:
        resp = requests.post(
            f"{llm_cfg['base_url']}/chat/completions",
            json=payload, headers=headers,
            timeout=llm_cfg.get("timeout", 30),
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"[post_process] 占位符LLM请求失败: {e}")
        return {}

    import json
    try:
        m = re.search(r"\{.*\}", content, re.S)
        if not m:
            raise ValueError("no JSON found")
        obj = json.loads(m.group(0))
        result = {}
        for ph in placeholders:
            v = obj.get(ph, "")
            if v is None:
                v = ""
            result[ph] = str(v)
        return result
    except Exception as e:
        print(f"[post_process] 占位符LLM响应解析失败: {e}")
        return {}


def _replace_placeholders_in_paragraph(p_el, mapping: dict) -> int:
    """在一段内跨 run 替换 {{XXX}}。简化策略：把段落文本拼成一个字符串替换后写回第一个 run，其余 run 清空文本。

    返回替换次数。
    """
    # 1. 收集整段文本与占位符位置
    raw_runs = p_el.findall(qn("w:r"))
    if not raw_runs:
        return 0

    full = "".join(t.text or "" for r in raw_runs for t in r.findall(qn("w:t")))
    if not PLACEHOLDER_RE.search(full):
        return 0

    def _do_sub(text: str) -> Tuple[str, int]:
        cnt = 0
        def _sub(m):
            nonlocal cnt
            key = m.group(1)
            cnt += 1
            return mapping.get(key, m.group(0))
        return PLACEHOLDER_RE.sub(_sub, text), cnt

    # 策略：逐 run 单独替换（简单替换，避免 run 属性丢失）；
    # 占位符跨 run 的情况兜底：把第一个保留 run 设为整段替换。
    total = 0
    any_replaced = False
    for r in raw_runs:
        t_elems = r.findall(qn("w:t"))
        if not t_elems:
            continue
        orig = "".join(t.text or "" for t in t_elems)
        if not PLACEHOLDER_RE.search(orig):
            continue
        new_txt, n = _do_sub(orig)
        if n > 0:
            # 把文本只写回第一个 w:t，其余 w:t 清空
            t_elems[0].text = new_txt
            for extra_t in t_elems[1:]:
                extra_t.text = ""
            total += n
            any_replaced = True

    if not any_replaced:
        # 跨 run 替换兜底：用 p_el 文本整体替换后写回第一个 w:t，其余 w:t 清空
        new_full, n = _do_sub(full)
        if n > 0:
            first_t = None
            for r in raw_runs:
                for t in r.findall(qn("w:t")):
                    if first_t is None:
                        first_t = t
                        t.text = new_full
                    else:
                        t.text = ""
            total = n
    return total


def replace_placeholders_with_llm(cfg: Config, doc: Document) -> dict:
    body_snippet, placeholders = collect_placeholders(doc)
    if not placeholders:
        return {"placeholders_found": 0, "replaced": 0, "keys": []}
    ph_list = sorted(placeholders)
    mapping = _llm_infer_placeholder_values(cfg, body_snippet, ph_list)
    replaced_count = 0

    def scan(paragraphs_iter):
        nonlocal replaced_count
        for p in paragraphs_iter:
            replaced_count += _replace_placeholders_in_paragraph(p._p, mapping)

    scan(doc.paragraphs)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                scan(cell.paragraphs)

    return {"placeholders_found": len(ph_list), "replaced": replaced_count,
            "keys": ph_list, "values": mapping}


# ========== 5. 更新术语表（EXPLANATION OF TERMS / 术语解释）==========

def _cell_text(cell) -> str:
    txt_parts = []
    for p in cell.paragraphs:
        for t in p._p.iter(qn("w:t")):
            txt_parts.append(t.text or "")
        txt_parts.append("\n")
    s = "".join(txt_parts).strip()
    s = s.replace("\x07", "")
    return s


def _find_glossary_table(doc: Document):
    """返回 (title_p, table, is_chinese)；找不到返回 (None, None, None)。
    
    支持模糊匹配：术语解释 / 术语表 / EXPLANATION OF TERMS / GLOSSARY 等。
    """
    body = doc.element.body
    title_p_el = None
    is_chinese: bool | None = None

    for p in body.findall(qn("w:p")):
        txt = _para_text(p).strip()
        if txt in ("术语解释", "术语表"):
            title_p_el = p
            is_chinese = True
            break
        if txt.upper() in ("EXPLANATION OF TERMS", "GLOSSARY", "TERMS"):
            title_p_el = p
            is_chinese = False
            break

    if title_p_el is None:
        return None, None, None

    # 找标题之后最近的 Table（通过 w:tbl 在 body 中顺序判断）
    body_children = list(body)
    title_idx = None
    for i, ch in enumerate(body_children):
        if ch is title_p_el:
            title_idx = i
            break
    if title_idx is None:
        return None, None, None

    best_tbl = None
    for i in range(title_idx + 1, len(body_children)):
        if body_children[i].tag == qn("w:tbl"):
            best_tbl = body_children[i]
            break

    if best_tbl is None:
        return None, None, None

    # 找到对应的 python-docx Table 对象
    for t in doc.tables:
        if t._tbl is best_tbl:
            return title_p_el, t, is_chinese
    return title_p_el, None, is_chinese


def _build_search_text(doc: Document, g_table: DocxTable) -> str:
    """搜索范围：从文档开头到术语表之前的所有正文段落+表格文本。
    
    不区分语种，中英文合并搜索，大小写不敏感。
    跳过 SectionName 标记行。
    """
    body = doc.element.body
    body_children = list(body)

    # 术语表自身在 body 中的位置
    g_tbl_el = g_table._tbl
    g_start_idx = body_children.index(g_tbl_el) if g_tbl_el in body_children else len(body_children)

    parts: List[str] = []
    for i in range(0, g_start_idx):
        ch = body_children[i]
        if ch.tag == qn("w:p"):
            txt = _para_text(ch)
            # 跳过 SectionName 标记行
            if SECTION_MARKER_RE.search(txt):
                continue
            parts.append(txt)
        elif ch.tag == qn("w:tbl"):
            # 找到对应的 python-docx Table 读文本
            for tt in doc.tables:
                if tt._tbl is ch:
                    for row in tt.rows:
                        for cell in row.cells:
                            parts.append(_cell_text(cell))
                    break
    combined = "\n".join(parts)
    # 大小写不敏感匹配
    return combined.lower()


def update_terms_table(doc: Document) -> dict:
    """移植 VBA 更新术语表 + 双语搜索 + 保留行变黑。

    搜索范围：从文档开头到术语表之前的所有内容（含 BodyText 正文）。
    匹配逻辑：术语表第一列（Acronym/术语）的值在搜索文本中查找，
    找到则保留行并变黑，未找到则删除该行。

    返回: {"found": bool, "kept": n, "deleted": n, "is_chinese": bool|None}
    """
    title_p_el, g_table, is_chinese = _find_glossary_table(doc)
    if title_p_el is None:
        return {"found": False, "kept": 0, "deleted": 0, "is_chinese": None}
    if g_table is None:
        return {"found": False, "kept": 0, "deleted": 0, "is_chinese": is_chinese}

    if len(g_table.rows) < 2:
        return {"found": True, "kept": 0, "deleted": 0, "is_chinese": is_chinese}

    search_text = _build_search_text(doc, g_table)

    deleted = 0
    kept = 0
    # 从后往前删，除表头
    for i in range(len(g_table.rows) - 1, 0, -1):
        row = g_table.rows[i]
        try:
            term_raw = _cell_text(row.cells[0])
        except Exception:
            continue
        term = (term_raw or "").strip()
        if not term:
            kept += 1
            continue
        # 不区分语种 & 大小写；中英文都在 search_text 里
        if term.lower() in search_text:
            kept += 1
            # 保留行字体变黑
            for cell in row.cells:
                for p in cell.paragraphs:
                    _set_run_color_black(p._p)
        else:
            try:
                row._tr.getparent().remove(row._tr)
                deleted += 1
            except Exception:
                pass

    return {"found": True, "kept": kept, "deleted": deleted, "is_chinese": is_chinese}


# ========== 主入口：一键执行后处理 ==========

def run_post_process(cfg: Config, merged_docx: str, output_path: str) -> str:
    """合并完成后，执行：删SectionName标记→表格文字→黑→占位符LLM→术语表→(BodyText范围内)图片/表格收尾。"""
    doc = Document(merged_docx)
    features = cfg.raw.get("features", {})

    # 先收集 BodyText 节范围（**必须在删 marker 之前**，否则 marker 已删找不到）
    from typeset_extra import format_images, format_tables
    bt_blocks = collect_bodytext_block_elements(doc, cfg.section_marker)

    # 4.1 删除所有 [SectionName: XXX] 标记行
    if features.get("remove_all_section_markers", True):
        st = remove_all_section_markers(doc)
        print(f"[post] 删除 SectionName 标记: {st['removed']} 行")

    # 4.2 占位符 {{XXX}} LLM 替换
    if features.get("replace_placeholders", True):
        st = replace_placeholders_with_llm(cfg, doc)
        print(f"[post] 占位符替换: 发现{st['placeholders_found']}个 ({st.get('keys')}) "
              f"实际替换 {st['replaced']} 处")

    # 5. 更新术语表（双语种）
    if features.get("update_terms_table", True):
        st = update_terms_table(doc)
        if st["found"]:
            lang = "中" if st["is_chinese"] else "EN"
            print(f"[post] 术语表更新: 语种={lang} "
                  f"保留 {st['kept']} 条，删除 {st['deleted']} 条")
        else:
            print("[post] 术语表更新: 未找到术语解释/EXPLANATION OF TERMS 标题或其下表格")

    # 4.3 所有表格文字颜色改黑（术语表已黑，再统一刷一遍不影响）
    if features.get("blacken_tables", True):
        st = blacken_all_table_text(doc)
        print(f"[post] 表格文字变黑: {st['tables']} 表 x {st['cells']} 单元格")

    # 1+2：图片 / 表格收尾（bt_blocks 仅 BodyText 节范围）
    if features.get("format_tables", True):
        t_stats = format_tables(doc, force_black_text=True)
        print(f"[post] 表格排版(收尾): 总{t_stats['total']} 处理{t_stats['processed']} "
              f"代码块{t_stats['code_block']} 跳过{t_stats['skipped']}")

    if features.get("format_images", True):
        is_cn = features.get("caption_chinese", True)
        i_stats = format_images(doc, is_chinese=is_cn, only_elements_in=bt_blocks)
        print(f"[post] 图片排版(收尾): 总{i_stats['total']} 处理{i_stats['processed']} "
              f"超范围跳过{i_stats['skipped_scope']} 过小跳过{i_stats['skipped_size']} 已样式跳过{i_stats['skipped_styled']}")

    import os
    out_dir = os.path.dirname(os.path.abspath(output_path)) or "."
    os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)
    print(f"[post] 后处理完成 -> {output_path}")
    return output_path
