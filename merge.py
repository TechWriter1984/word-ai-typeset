"""步骤3：基于模板合并，替换 [SectionName: BodyText] 节内容。

逻辑（与原 VBA 安全粘贴正文 等价）：
1. 复制模板文件为输出文件（保留封面、页眉页脚、域、目录等）
2. 在输出文件中找到包含 [SectionName: BodyText] 的标记段落
3. 清空标记之后、文档末尾分节符之前的所有占位内容
4. 把排版后 docx 的正文（段落 + 表格，按顺序）追加到标记之后
5. 段落/表格样式按"名称"重新映射到模板的对应样式
   （避免两个 docx 的 styleId 不一致导致样式丢失）
"""
from __future__ import annotations

import copy
import os
import shutil
from typing import Optional

from docx import Document
from docx.oxml.ns import qn

from config import Config


def _find_marker_paragraph(doc: Document, marker: str):
    """返回 body 中包含 marker 文本的 <w:p> 元素，找不到返回 None。"""
    body = doc.element.body
    for p in body.iter(qn("w:p")):
        text = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if marker in text:
            return p
    return None


def _style_name_by_id(styles_element, style_id: str) -> Optional[str]:
    """在 styles.xml 中按 styleId 反查样式名（w:name w:val）。"""
    for s in styles_element.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == style_id:
            name_el = s.find(qn("w:name"))
            if name_el is not None:
                return name_el.get(qn("w:val"))
    return None


def _style_id_by_name(styles_element, name: str) -> Optional[str]:
    """在 styles.xml 中按样式名反查 styleId。"""
    for s in styles_element.findall(qn("w:style")):
        name_el = s.find(qn("w:name"))
        if name_el is not None and name_el.get(qn("w:val")) == name:
            return s.get(qn("w:styleId"))
    return None


def _remap_paragraph_style(p_el, src_styles, tpl_styles) -> None:
    """把段落的 pStyle w:val 从 src 的 styleId 改写为 tpl 的 styleId（按名称匹配）。"""
    pStyle = p_el.find(qn("w:pStyle"))
    if pStyle is None:
        return
    src_id = pStyle.get(qn("w:val"))
    if not src_id:
        return
    name = _style_name_by_id(src_styles, src_id)
    if not name:
        return
    tpl_id = _style_id_by_name(tpl_styles, name)
    if tpl_id and tpl_id != src_id:
        pStyle.set(qn("w:val"), tpl_id)
    elif tpl_id is None:
        # 模板没有同名样式：移除 pStyle 引用，回退到默认 Normal
        p_el.remove(pStyle)


def merge_template(cfg: Config, src_docx: str, out_path: str) -> str:
    """以模板为基底新建文件，把 src_docx 内容替换到 BodyText 节。"""
    template_path = cfg.template_docx
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"模板不存在: {template_path}")
    if not os.path.exists(src_docx):
        raise FileNotFoundError(f"源文件不存在: {src_docx}")

    out_dir = os.path.dirname(os.path.abspath(out_path)) or "."
    os.makedirs(out_dir, exist_ok=True)

    # 1. 复制模板为输出文件（保留模板全部结构：封面、页眉页脚、域、目录等）
    shutil.copyfile(template_path, out_path)

    tpl_doc = Document(out_path)
    src_doc = Document(src_docx)

    marker_p = _find_marker_paragraph(tpl_doc, cfg.section_marker)
    if marker_p is None:
        raise RuntimeError(f"模板中未找到标记段落: [SectionName: {cfg.section_marker}]")

    body = tpl_doc.element.body
    sect_pr = body.find(qn("w:sectPr"))

    # 2. 删除 marker 段本身 + 之后到 sectPr 之前的占位内容
    #    （用户要求删除 [SectionName: BodyText] 这一行）
    #    删除前先记录 marker 的前一个兄弟作插入锚点，保证源内容插到 BodyText 节原位置
    prev_sibling = marker_p.getprevious()
    to_remove = [marker_p]
    if cfg.clear_placeholder:
        started = False
        for child in list(body):
            if not started:
                if child is marker_p:
                    started = True
                continue
            if child is sect_pr:
                break
            to_remove.append(child)
    for c in to_remove:
        body.remove(c)
    print(f"[merge] 已删除 marker + 占位内容 {len(to_remove)} 个块")

    # 3. 把 src body 的段落/表格按顺序插入到 marker 原位置（prev_sibling 之后）
    src_body = src_doc.element.body
    src_styles = src_doc.part.styles.element
    tpl_styles = tpl_doc.part.styles.element

    anchor = prev_sibling  # 可能为 None（marker 原是第一个），此时插到 body 最前
    inserted = 0
    for child in list(src_body):
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag not in (qn("w:p"), qn("w:tbl")):
            continue
        new_el = copy.deepcopy(child)
        if anchor is None:
            body.insert(0, new_el)
        else:
            anchor.addnext(new_el)
        anchor = new_el
        inserted += 1
        if new_el.tag == qn("w:p"):
            _remap_paragraph_style(new_el, src_styles, tpl_styles)

    tpl_doc.save(out_path)
    print(f"[merge] 完成 -> {out_path}  (插入 {inserted} 个块)")
    return out_path
