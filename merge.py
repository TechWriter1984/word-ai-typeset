"""步骤3：基于模板合并，替换 [SectionName: BodyText] 节内容。

逻辑（与原 VBA 安全粘贴正文 等价）：
1. 复制模板文件为输出文件（保留封面、页眉页脚、域、目录等）
2. 在输出文件中找到包含 [SectionName: BodyText] 的标记段落
3. 清空标记之后、文档末尾分节符之前的所有占位内容
4. 把排版后 docx 的正文（段落 + 表格，按顺序）追加到标记之后
5. 段落/表格样式按"名称"重新映射到模板的对应样式
6. 复制源文档的图片/图表 parts，并重建 r:embed 引用
   （否则 deepcopy 后图片会显示"无法显示该图片"）
"""
from __future__ import annotations

import copy
import os
import shutil
import tempfile
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

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
        p_el.remove(pStyle)


def _collect_src_rid_to_part(src_doc) -> dict:
    """收集源 document.xml 所有 rel -> image/chart/ole part。"""
    mapping = {}
    for rid, rel in src_doc.part.rels.items():
        rt = rel.reltype
        if (rt == RT.IMAGE
                or "/relationships/chart" in rt
                or "/relationships/oleObject" in rt
                or "/relationships/package" in rt):
            try:
                mapping[rid] = rel.target_part
            except Exception:
                pass
    return mapping


def _clone_image_part_to_tgt(src_part, tgt_doc_part) -> str:
    """把 src_part 的 blob 复制到 tgt_doc_part，返回新 rId。

    使用 python-docx 内置的 get_or_add_image_part 机制，
    确保图片 part 正确注册到包中（Content_Types + rels），
    并利用 SHA1 哈希自动去重——相同图片只存一份。
    """
    try:
        blob = src_part.blob
        ext = os.path.splitext(src_part.partname)[1].lower() or ".png"
    except Exception:
        return ""

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(blob)
            tmp_path = tmp.name
        # get_or_add_image_part 内部用 SHA1 去重，
        # 返回的 rId 可直接用于 r:embed
        return tgt_doc_part.get_or_add_image_part(tmp_path)
    except Exception:
        return ""
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


def _rewrite_embeds_in_element(el, src_rid_to_part, tgt_doc_part, rid_cache: dict):
    """递归 el 下所有 r:embed / r:link / r:id 属性，把 src rId 换成 tgt 新 rId。"""
    for elem in el.iter():
        for attr_name in list(elem.attrib.keys()):
            if attr_name in (qn("r:embed"), qn("r:link"), qn("r:id")):
                old_rid = elem.attrib[attr_name]
                if old_rid in rid_cache:
                    new_rid = rid_cache[old_rid]
                elif old_rid in src_rid_to_part:
                    new_rid = _clone_image_part_to_tgt(src_rid_to_part[old_rid], tgt_doc_part)
                    rid_cache[old_rid] = new_rid
                else:
                    continue
                if new_rid:
                    elem.attrib[attr_name] = new_rid


def merge_template(cfg: Config, src_docx: str, out_path: str) -> str:
    """以模板为基底新建文件，把 src_docx 内容替换到 BodyText 节。"""
    template_path = cfg.template_docx
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"模板不存在: {template_path}")
    if not os.path.exists(src_docx):
        raise FileNotFoundError(f"源文件不存在: {src_docx}")

    out_dir = os.path.dirname(os.path.abspath(out_path)) or "."
    os.makedirs(out_dir, exist_ok=True)

    shutil.copyfile(template_path, out_path)

    tpl_doc = Document(out_path)
    src_doc = Document(src_docx)

    marker_p = _find_marker_paragraph(tpl_doc, cfg.section_marker)
    if marker_p is None:
        raise RuntimeError(f"模板中未找到标记段落: [SectionName: {cfg.section_marker}]")

    body = tpl_doc.element.body
    sect_pr = body.find(qn("w:sectPr"))

    # marker 段本身暂不删除，留到 post_process.remove_all_section_markers 阶段统一删
    # （这样 post_process 才能用 marker 定位 BodyText 节的起始边界）
    to_remove = []
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
    print(f"[merge] 已删除占位内容 {len(to_remove)} 个块（BodyText marker 保留到后处理阶段再删）")

    src_rid_to_part = _collect_src_rid_to_part(src_doc)
    rid_cache = {}

    src_body = src_doc.element.body
    src_styles = src_doc.part.styles.element
    tpl_styles = tpl_doc.part.styles.element

    # 插入位置：BodyText marker 之后（marker_p 与 sectPr 之间），
    # 这样 post_process 可用 marker 定位 BodyText 节范围，只处理源内容。
    anchor = marker_p
    inserted = 0
    for child in list(src_body):
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag not in (qn("w:p"), qn("w:tbl")):
            continue
        new_el = copy.deepcopy(child)

        if src_rid_to_part:
            _rewrite_embeds_in_element(new_el, src_rid_to_part, tpl_doc.part, rid_cache)

        # addnext: 加到锚点之后；锚点递进，保证插入顺序与源文档一致
        anchor.addnext(new_el)
        anchor = new_el
        inserted += 1
        if new_el.tag == qn("w:p"):
            _remap_paragraph_style(new_el, src_styles, tpl_styles)

    if rid_cache:
        print(f"[merge] 图片/资源 parts: 源 {len(src_rid_to_part)} 个，复制+重建 {len(rid_cache)} 条引用")

    tpl_doc.save(out_path)
    print(f"[merge] 完成 -> {out_path}  (插入 {inserted} 个块)")
    return out_path
