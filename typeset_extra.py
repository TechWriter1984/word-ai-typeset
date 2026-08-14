"""typeset_extra.py：表格排版 / 图片排版 / 清理标题多余编号。

对齐 VBA 宏：批量表格排版 / 批量图片排版 / 清理标题多余编号。
全部用 python-docx + oxml 实现，不依赖 Word COM。

新增：
- format_images 支持传入「允许处理的 XML 元素范围」（BodyText 节起止）
  非范围内的 inline_shape 一律跳过（如封面 Logo、SYMBOL 表 Icon）
- format_tables 给表头行加重复标题行（w:tblHeader）
  并确保对齐：表头水平居中（段落 jc 覆盖样式默认），内容中部左对齐
"""
from __future__ import annotations

from docx import Document
from docx.oxml.ns import qn

STYLE_TABLE_HEADER = "_表格_表头标题行_hs"
STYLE_TABLE_BODY = "_表格_正文格式_hs"
STYLE_CODE = "_代码字体_hs"
STYLE_IMAGE = "_图_hs"
STYLE_CAPTION = "_图_题注_hs"

GRAY_FILL = "D1D1D1"
IMG_BORDER_COLOR = "4472C4"

# Logo 参考尺寸：汉朔 Logo = 1.51cm x 6.84cm（1cm=360000 EMU）
LOGO_W_EMU = int(1.51 * 360000)
LOGO_H_EMU = int(6.84 * 360000)
# 最小面积阈值（比 Logo 大 1.3 倍以上才处理）
MIN_AREA_RATIO = 1.3


# ========== 通用 oxml 辅助 ==========

def _style_id_to_name(doc: Document, sid: str) -> str:
    if not sid:
        return ""
    for st in doc.part.styles.element.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == sid:
            nm = st.find(qn("w:name"))
            return nm.get(qn("w:val")) if nm is not None else ""
    return ""


def _para_style_id(p_el) -> str:
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return ""
    pStyle = pPr.find(qn("w:pStyle"))
    return pStyle.get(qn("w:val")) if pStyle is not None else ""


def _para_style_name(doc, p_el) -> str:
    return _style_id_to_name(doc, _para_style_id(p_el))


def _set_para_align_el(p_el, val="center"):
    """设段落水平对齐，放在 pPr 第一个位置以确保比样式默认优先级高。"""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = p_el.makeelement(qn("w:pPr"), {})
        p_el.insert(0, pPr)
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = pPr.makeelement(qn("w:jc"), {})
        pPr.insert(0, jc)
    else:
        pPr.remove(jc)
        jc_new = pPr.makeelement(qn("w:jc"), {})
        pPr.insert(0, jc_new)
        jc = jc_new
    jc.set(qn("w:val"), val)


def _set_para_style_by_name(doc, p_el, name):
    sid = ""
    for st in doc.part.styles.element.findall(qn("w:style")):
        nm = st.find(qn("w:name"))
        if nm is not None and nm.get(qn("w:val")) == name:
            sid = st.get(qn("w:styleId")) or ""
            break
    if not sid:
        return False
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = p_el.makeelement(qn("w:pPr"), {})
        p_el.insert(0, pPr)
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = pPr.makeelement(qn("w:pStyle"), {})
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), sid)
    return True


def _set_run_color_in_paragraph(p_el, color_hex: str):
    """把段落内所有 w:r 的字体颜色强制设为 color_hex（如 "000000"）。"""
    for r in p_el.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = r.makeelement(qn("w:rPr"), {})
            r.insert(0, rPr)
        color = rPr.find(qn("w:color"))
        if color is None:
            color = rPr.makeelement(qn("w:color"), {})
            rPr.append(color)
        color.set(qn("w:val"), color_hex)


# ========== 1. 批量表格排版 ==========

def _set_cell_shading(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if fill_hex == "auto":
        if shd is not None:
            tcPr.remove(shd)
        return
    if shd is None:
        shd = tcPr.makeelement(qn("w:shd"), {})
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)


def _set_cell_v_align(cell, val="center"):
    tcPr = cell._tc.get_or_add_tcPr()
    vA = tcPr.find(qn("w:vAlign"))
    if vA is None:
        vA = tcPr.makeelement(qn("w:vAlign"), {})
        tcPr.append(vA)
    vA.set(qn("w:val"), val)


def _set_row_repeat_header(row):
    """Word「重复标题行」：<w:trPr><w:tblHeader/></w:trPr>。"""
    tr = row._tr
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = tr.makeelement(qn("w:trPr"), {})
        tr.insert(0, trPr)
    tblHeader = trPr.find(qn("w:tblHeader"))
    if tblHeader is None:
        tblHeader = trPr.makeelement(qn("w:tblHeader"), {})
        trPr.append(tblHeader)


def _set_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = tbl.makeelement(qn("w:tblPr"), {})
        tbl.insert(0, tblPr)
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = tblPr.makeelement(qn("w:tblBorders"), {})
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = borders.find(qn(f"w:{edge}"))
        if b is None:
            b = borders.makeelement(qn(f"w:{edge}"), {})
            borders.append(b)
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")


def _set_table_autofit(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = tbl.makeelement(qn("w:tblPr"), {})
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = tblPr.makeelement(qn("w:tblW"), {})
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"), "5000")
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = tblPr.makeelement(qn("w:tblLayout"), {})
        tblPr.append(layout)
    layout.set(qn("w:type"), "autofit")


def _apply_style_to_cell_paragraphs(doc, cell, name):
    for p in cell.paragraphs:
        _set_para_style_by_name(doc, p._p, name)


def format_tables(doc: Document, force_black_text: bool = False) -> dict:
    """批量表格排版（对齐 VBA 批量表格排版）。

    参数:
        force_black_text: True 时把除代码块外所有单元格内文字颜色强制设黑
    """
    stats = {"total": 0, "processed": 0, "skipped": 0, "code_block": 0}
    for tbl in doc.tables:
        stats["total"] += 1
        nrows = len(tbl.rows)
        ncols = len(tbl.columns)
        if nrows == 0:
            continue

        if nrows == 1 and ncols == 1:
            _apply_style_to_cell_paragraphs(doc, tbl.cell(0, 0), STYLE_CODE)
            _set_table_autofit(tbl)
            if force_black_text:
                for p in tbl.cell(0, 0).paragraphs:
                    _set_run_color_in_paragraph(p._p, "000000")
            stats["code_block"] += 1
            continue

        first_name = _para_style_name(doc, tbl.cell(0, 0).paragraphs[0]._p) if tbl.cell(0, 0).paragraphs else ""
        if first_name == STYLE_TABLE_HEADER and nrows >= 2:
            second_name = _para_style_name(doc, tbl.cell(1, 0).paragraphs[0]._p) if tbl.cell(1, 0).paragraphs else ""
            if second_name == STYLE_TABLE_BODY:
                stats["skipped"] += 1
                if force_black_text:
                    for r_idx in range(1, nrows):
                        for cell in tbl.rows[r_idx].cells:
                            for p in cell.paragraphs:
                                _set_run_color_in_paragraph(p._p, "000000")
                continue

        _set_table_autofit(tbl)
        _set_table_borders(tbl)
        for ri, row in enumerate(tbl.rows):
            if ri == 0:
                _set_row_repeat_header(row)
                for cell in row.cells:
                    _apply_style_to_cell_paragraphs(doc, cell, STYLE_TABLE_HEADER)
                    _set_cell_v_align(cell, "center")
                    for p in cell.paragraphs:
                        _set_para_align_el(p._p, "center")
                        if force_black_text:
                            _set_run_color_in_paragraph(p._p, "000000")
                    _set_cell_shading(cell, GRAY_FILL)
            else:
                for cell in row.cells:
                    _apply_style_to_cell_paragraphs(doc, cell, STYLE_TABLE_BODY)
                    _set_cell_v_align(cell, "center")
                    for p in cell.paragraphs:
                        _set_para_align_el(p._p, "left")
                        if force_black_text:
                            _set_run_color_in_paragraph(p._p, "000000")
                    _set_cell_shading(cell, "auto")
        stats["processed"] += 1
    return stats


# ========== 2. 批量图片排版 ==========

def _find_body_level_ancestor(p_el):
    """从 p_el 向上遍历，找到 body 的直接子级（w:p 或 w:tbl）。
    
    对于表格内的图片，父级链是 p->tc->tr->tbl->body，
    循环需要穿透所有中间层，直到父级是 w:body 为止。
    返回 body 的直接子级元素。
    """
    top_el = p_el
    while top_el is not None:
        parent = top_el.getparent()
        if parent is None:
            break
        if parent.tag == qn("w:body"):
            # top_el 现在是 body 的直接子级
            return top_el
        top_el = parent
    # 兜底：返回原始 p_el
    return p_el


def _shape_paragraph_el(shape):
    el = shape._inline
    while el is not None and el.tag != qn("w:p"):
        el = el.getparent()
    return el


def _set_image_border(shape):
    inline = shape._inline
    graphic = inline.find(qn("a:graphic"))
    if graphic is None:
        return
    graphicData = graphic.find(qn("a:graphicData"))
    if graphicData is None:
        return
    pic = graphicData.find(qn("pic:pic"))
    if pic is None:
        return
    picPr = pic.find(qn("pic:picPr"))
    if picPr is None:
        picPr = pic.makeelement(qn("pic:picPr"), {})
        pic.insert(0, picPr)
    picBdr = picPr.find(qn("pic:picBdr"))
    if picBdr is None:
        picBdr = picPr.makeelement(qn("pic:picBdr"), {})
        picPr.append(picBdr)
    for edge in ("top", "left", "bottom", "right"):
        b = picBdr.find(qn(f"a:{edge}"))
        if b is None:
            b = picBdr.makeelement(qn(f"a:{edge}"), {})
            picBdr.append(b)
        b.set("w", "12700")
        b.set("cap", "flat")
        b.set("cmpd", "sng")
        b.set("algn", "ctr")
        solidFill = b.find(qn("a:solidFill"))
        if solidFill is None:
            solidFill = b.makeelement(qn("a:solidFill"), {})
            b.append(solidFill)
        srgb = solidFill.find(qn("a:srgbClr"))
        if srgb is None:
            srgb = solidFill.makeelement(qn("a:srgbClr"), {})
            solidFill.append(srgb)
        srgb.set("val", IMG_BORDER_COLOR)


def _add_field(run_el, instr):
    fldBegin = run_el.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "begin"})
    instrText = run_el.makeelement(qn("w:instrText"), {qn("xml:space"): "preserve"})
    instrText.text = " " + instr + " "
    fldSep = run_el.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "separate"})
    placeholder = run_el.makeelement(qn("w:t"), {})
    placeholder.text = "1"
    fldEnd = run_el.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "end"})
    run_el.append(fldBegin)
    run_el.append(instrText)
    run_el.append(fldSep)
    run_el.append(placeholder)
    run_el.append(fldEnd)


def _build_caption_paragraph(doc, is_chinese: bool):
    label = "图" if is_chinese else "Figure"
    seq_name = "图" if is_chinese else "Figure"

    p_el = doc.element.body.makeelement(qn("w:p"), {})
    pPr = p_el.makeelement(qn("w:pPr"), {})
    p_el.append(pPr)
    jc = pPr.makeelement(qn("w:jc"), {qn("w:val"): "center"})
    pPr.append(jc)
    _set_para_style_by_name(doc, p_el, STYLE_CAPTION)

    r1 = p_el.makeelement(qn("w:r"), {})
    t1 = r1.makeelement(qn("w:t"), {})
    t1.text = label + " "
    r1.append(t1)
    p_el.append(r1)

    r2 = p_el.makeelement(qn("w:r"), {})
    p_el.append(r2)
    _add_field(r2, "STYLEREF 1 \\s")

    r3 = p_el.makeelement(qn("w:r"), {})
    t3 = r3.makeelement(qn("w:t"), {})
    t3.text = "-"
    r3.append(t3)
    p_el.append(r3)

    r4 = p_el.makeelement(qn("w:r"), {})
    p_el.append(r4)
    _add_field(r4, f"SEQ {seq_name} \\* ARABIC \\s1")

    return p_el


def format_images(doc: Document, is_chinese: bool = True,
                  only_elements_in: list | None = None,
                  size_filter_ratio: float = MIN_AREA_RATIO) -> dict:
    """批量图片排版。

    参数:
        only_elements_in: 若给出（body 级别的子元素列表），
            只处理图片所在段落属于该列表中的元素（用于只处理 BodyText 节）。
            若为 None 则处理所有 inline_shapes。
        size_filter_ratio: 宽或高 任一维度 > Logo * ratio 的才处理
            （ratio=1.3 表示宽>1.96cm 或 高>8.89cm 才处理，排除小图标）
    """
    stats = {"total": 0, "processed": 0, "skipped_scope": 0, "skipped_size": 0, "skipped_styled": 0}

    allowed_set = set(id(x) for x in only_elements_in) if only_elements_in else None

    shapes = list(doc.inline_shapes)
    for shape in shapes:
        stats["total"] += 1
        p_el = _shape_paragraph_el(shape)
        if p_el is None:
            stats["skipped_scope"] += 1
            continue

        # 范围过滤：只处理 BodyText 节里的段落/表格
        if allowed_set is not None:
            # 穿透完整父级链找到 body 的直接子级（w:p 或 w:tbl）
            top_el = _find_body_level_ancestor(p_el)
            if id(top_el) not in allowed_set:
                stats["skipped_scope"] += 1
                continue

        cur_name = _para_style_name(doc, p_el)
        if cur_name == STYLE_IMAGE:
            stats["skipped_styled"] += 1
            continue

        w = shape.width or 0
        h = shape.height or 0
        # 尺寸过滤：至少一维明显大于 Logo（防止处理封面 Logo / 表内小图标）
        if not (w > LOGO_W_EMU * size_filter_ratio or h > LOGO_H_EMU * size_filter_ratio):
            stats["skipped_size"] += 1
            continue

        _set_para_align_el(p_el, "center")
        _set_image_border(shape)
        _set_para_style_by_name(doc, p_el, STYLE_IMAGE)
        cap_p = _build_caption_paragraph(doc, is_chinese)
        p_el.addnext(cap_p)
        stats["processed"] += 1
    return stats


# ========== 3. 清理标题多余编号 ==========

def _is_chinese_digit(ch: str) -> bool:
    return ch in "零一二三四五六七八九十百千两〇"


def _leading_number_len(s: str) -> int:
    if not s:
        return 0
    i = 0
    n = len(s)
    while i < n and s[i] in (" ", "\t", "　"):
        i += 1
    if i < n and s[i] == "第":
        i += 1
    num_found = False
    while i < n:
        ch = s[i]
        if ch.isdigit() or _is_chinese_digit(ch):
            num_found = True
            i += 1
        elif ch in (".", "．") and num_found:
            i += 1
        else:
            break
    if not num_found:
        return 0
    if i < n and s[i] in ("章", "节"):
        i += 1
    sep_found = False
    while i < n:
        ch = s[i]
        if ch in ".)）、,，:：． 　\t":
            i += 1
            sep_found = True
        else:
            break
    if not sep_found:
        return 0
    if i > n:
        return 0
    return i


def _delete_leading_chars(para, n: int) -> None:
    deleted = 0
    for run in list(para.runs):
        if deleted >= n:
            break
        t = run.text
        if not t:
            continue
        take = min(len(t), n - deleted)
        run.text = t[take:]
        deleted += take
        if not run.text:
            run._r.getparent().remove(run._r)


def clean_heading_numbers(doc: Document) -> dict:
    stats = {"cleaned": 0, "skipped": 0, "not_heading": 0}
    for para in doc.paragraphs:
        p_el = para._p
        name = _para_style_name(doc, p_el)
        if not (name.startswith("heading ") and name[8:].isdigit()):
            stats["not_heading"] += 1
            continue
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            stats["skipped"] += 1
            continue
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            stats["skipped"] += 1
            continue
        text = para.text or ""
        if not text:
            stats["skipped"] += 1
            continue
        n = _leading_number_len(text)
        if n > 0:
            _delete_leading_chars(para, n)
            stats["cleaned"] += 1
        else:
            stats["skipped"] += 1
    return stats
