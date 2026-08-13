"""步骤2：LLM 分类 + 自动排版。

输入：用户前处理后的 docx 路径
行为：调用大模型对每段做样式分类，按 config.style_map 映射到
      Word 自定义样式名并应用。
输出：排版后的 docx（保存到 cfg.output_dir，文件名加后缀 _typeset）
"""
from __future__ import annotations

import os
from typing import List, Tuple

import copy

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from config import Config
from llm import classify_paragraphs
from typeset_extra import format_tables, format_images, clean_heading_numbers


def _import_styles_from_template(src_doc: Document, tpl_path: str, needed_names) -> int:
    """从模板把 needed_names 里的样式定义复制到源 docx（已存在的不覆盖）。

    等价 VBA 的 Application.OrganizerCopy。只复制 <w:style> 定义，
    列表编号定义（numbering.xml）不复制——merge 到模板后由模板 numbering 接管。
    """
    if not os.path.exists(tpl_path):
        return 0
    tpl_doc = Document(tpl_path)
    src_styles_el = src_doc.part.styles.element
    tpl_styles_el = tpl_doc.part.styles.element

    existing = set()
    for st in src_styles_el.findall(qn("w:style")):
        nm = st.find(qn("w:name"))
        if nm is not None:
            existing.add(nm.get(qn("w:val")))

    imported = 0
    for st in tpl_styles_el.findall(qn("w:style")):
        nm = st.find(qn("w:name"))
        if nm is None:
            continue
        name = nm.get(qn("w:val"))
        if name not in needed_names or name in existing:
            continue
        src_styles_el.append(copy.deepcopy(st))
        imported += 1
    return imported


def _collect_paragraphs(doc: Document, skip_blank: bool, max_len: int) -> Tuple[List[str], List[int]]:
    """收集需要分类的段落文本及其在 doc.paragraphs 中的下标。

    注：仅遍历顶层段落，不含表格内段落（表格样式由后续合并步骤保留）。
    """
    items: List[str] = []
    indices: List[int] = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if skip_blank and not text:
            continue
        items.append(text[:max_len])
        indices.append(idx)
    return items, indices


def typeset_doc(cfg: Config, input_path: str) -> str:
    """对 input_path 进行 LLM 分类并应用样式，返回输出路径。"""
    os.makedirs(cfg.output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(cfg.output_dir, f"{base}_typeset.docx")

    doc = Document(input_path)

    # 从模板导入缺失的自定义样式定义（等价 VBA OrganizerCopy）
    needed = set(cfg.style_map.values())
    imported = _import_styles_from_template(doc, cfg.template_docx, needed)
    if imported:
        print(f"[typeset] 从模板导入 {imported} 个样式定义")

    items, indices = _collect_paragraphs(
        doc,
        skip_blank=cfg.typeset.get("skip_blank", True),
        max_len=cfg.typeset.get("max_text_len", 500),
    )
    print(f"[typeset] 共 {len(items)} 段待分类")

    labels = classify_paragraphs(cfg, items)
    style_map = cfg.style_map
    missing = set()
    applied = 0

    # 先建一次「样式名 -> Styles 对象」缓存，避免每段重复查找
    style_cache = {}
    for name in set(style_map.values()):
        try:
            style_cache[name] = doc.styles[name]
        except KeyError:
            style_cache[name] = None

    for label, idx in zip(labels, indices):
        target = style_map.get(label)
        if not target:
            missing.add(label)
            continue
        sty = style_cache.get(target)
        if sty is None:
            missing.add(target)
            continue
        para: Paragraph = doc.paragraphs[idx]
        try:
            para.style = sty
            applied += 1
        except Exception as e:
            missing.add(f"{target}({type(e).__name__})")

    # ---- 增强排版（对齐 VBA 三宏）----
    features = cfg.raw.get("features", {})
    if features.get("format_tables", True):
        t_stats = format_tables(doc)
        print(f"[typeset] 表格排版: 总{t_stats['total']} 处理{t_stats['processed']} "
              f"代码块{t_stats['code_block']} 跳过{t_stats['skipped']}")
    if features.get("format_images", True):
        is_cn = features.get("caption_chinese", True)
        i_stats = format_images(doc, is_chinese=is_cn)
        print(f"[typeset] 图片排版: 总{i_stats['total']} 处理{i_stats['processed']} 跳过{i_stats['skipped']}")
    if features.get("clean_heading_numbers", True):
        c_stats = clean_heading_numbers(doc)
        print(f"[typeset] 清理标题编号: 清理{c_stats['cleaned']} 跳过{c_stats['skipped']} 非标题{c_stats['not_heading']}")

    doc.save(out_path)
    print(f"[typeset] 完成 -> {out_path}  (应用 {applied}/{len(labels)} 段)")
    if missing:
        print(f"[typeset] 注意：以下标签/样式未命中：{sorted(missing)}")
        # 诊断：列出源 docx 实际可用的目标样式名
        from docx.oxml.ns import qn
        avail = []
        for s in doc.part.styles.element.findall(qn("w:style")):
            nm = s.find(qn("w:name"))
            if nm is not None:
                v = nm.get(qn("w:val")) or ""
                if v.startswith("_") and v.endswith("_hs") or v.startswith("heading"):
                    avail.append(v)
        print(f"[typeset] 源 docx 可用目标样式（前 30）: {sorted(set(avail))[:30]}")
    return out_path
