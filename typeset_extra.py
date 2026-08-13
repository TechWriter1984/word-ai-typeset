"""typeset_extra.py：表格排版 / 图片排版 / 清理标题多余编号。

对齐 VBA 宏：批量表格排版 / 批量图片排版 / 清理标题多余编号。
全部用 python-docx + oxml 实现，不依赖 Word COM。
"""
from __future__ import annotations

from docx import Document
from docx.oxml.ns import qn

# ========== 样式名常量（与 config.yaml style_map 中的模板真实样式名一致）==========

STYLE_TABLE_HEADER = "_表格_表头标题行_hs"
STYLE_TABLE_BODY = "_表格_正文格式_hs"
STYLE_CODE = "_代码字体_hs"
STYLE_IMAGE = "_图_hs"
STYLE_CAPTION = "_图_题注_hs"

GRAY_FILL = "D1D1D1"        # RGB(209,209,209)
IMG_BORDER_COLOR = "4472C4"  # RGB(68,114,196)

# 图片尺寸阈值：1.51cm × 6.84cm（1cm=360000 EMU）
MIN_W_EMU = int(1.51 * 360000)
MIN_H_EMU = int(6.84 * 360000)


# ========== 通用 oxml 辅助 ==========

def _style_id_to_name(doc: Document, sid: str) -> str:
    """按 styleId 反查样式 name（w:name w:val）。"""
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
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = p_el.makeelement(qn("w:pPr"), {})
        p_el.insert(0, pPr)
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = pPr.makeelement(qn("w:jc"), {})
        pPr.append(jc)
    jc.set(qn("w:val"), val)


def _set_para_style_by_name(doc, p_el, name):
    """按样式名查 styleId 后写入 pStyle。"""
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
        b.set(qn("w:sz"), "4")  # 0.5pt (eighth-pt)
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
    tblW.set(qn("w:w"), "5000")  # 100%
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = tblPr.makeelement(qn("w:tblLayout"), {})
        tblPr.append(layout)
    layout.set(qn("w:type"), "autofit")


def _apply_style_to_cell_paragraphs(doc, cell, name):
    for p in cell.paragraphs:
        _set_para_style_by_name(doc, p._p, name)


def format_tables(doc: Document) -> dict:
    """批量表格排版（对齐 VBA 批量表格排版）。"""
    stats = {"total": 0, "processed": 0, "skipped": 0, "code_block": 0}
    for tbl in doc.tables:
        stats["total"] += 1
        nrows = len(tbl.rows)
        ncols = len(tbl.columns)

        # 1. 代码块：1行1列
        if nrows == 1 and ncols == 1:
            _apply_style_to_cell_paragraphs(doc, tbl.cell(0, 0), STYLE_CODE)
            _set_table_autofit(tbl)
            stats["code_block"] += 1
            continue

        # 2. 跳过已完成排版的表格
        first_name = _para_style_name(doc, tbl.cell(0, 0).paragraphs[0]._p) if tbl.cell(0, 0).paragraphs else ""
        if first_name == STYLE_TABLE_HEADER:
            if nrows == 1:
                stats["skipped"] += 1
                continue
            second_name = _para_style_name(doc, tbl.cell(1, 0).paragraphs[0]._p) if tbl.cell(1, 0).paragraphs else ""
            if second_name == STYLE_TABLE_BODY:
                stats["skipped"] += 1
                continue

        # 3. 处理普通表格
        _set_table_autofit(tbl)
        _set_table_borders(tbl)
        for ri, row in enumerate(tbl.rows):
            if ri == 0:
                # 表头：样式 + 水平居中 + 垂直居中 + 灰底
                for cell in row.cells:
                    _apply_style_to_cell_paragraphs(doc, cell, STYLE_TABLE_HEADER)
                    _set_cell_v_align(cell, "center")
                    for p in cell.paragraphs:
                        _set_para_align_el(p._p, "center")
                    _set_cell_shading(cell, GRAY_FILL)
            else:
                # 正文行：样式 + 中部左对齐 + 清背景
                for cell in row.cells:
                    _apply_style_to_cell_paragraphs(doc, cell, STYLE_TABLE_BODY)
                    _set_cell_v_align(cell, "center")
                    for p in cell.paragraphs:
                        _set_para_align_el(p._p, "left")
                    _set_cell_shading(cell, "auto")
        stats["processed"] += 1
    return stats


# ========== 2. 批量图片排版 ==========

def _shape_paragraph_el(shape):
    """返回 inline_shape 所在 <w:p> 元素。"""
    el = shape._inline
    while el is not None and el.tag != qn("w:p"):
        el = el.getparent()
    return el


def _set_image_border(shape):
    """给图片加 1pt 边框 RGB(68,114,196)。

    操作 drawing > graphic > graphicData > pic:pic > pic:picPr > pic:picBdr
    """
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
        b.set("w", "12700")  # 1pt = 12700 EMU
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
    """在 run 元素内追加域：begin / instrText / separate / end。"""
    nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
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
    r"""构造题注段落：<图/Figure> <STYLEREF 1 \s> - <SEQ 图 \* ARABIC \s1>。"""
    label = "图" if is_chinese else "Figure"
    seq_name = "图" if is_chinese else "Figure"

    p_el = doc.element.body.makeelement(qn("w:p"), {})
    pPr = p_el.makeelement(qn("w:pPr"), {})
    p_el.append(pPr)
    # 居中
    jc = pPr.makeelement(qn("w:jc"), {qn("w:val"): "center"})
    pPr.append(jc)
    # 样式
    _set_para_style_by_name(doc, p_el, STYLE_CAPTION)

    # run1: "图 "
    r1 = p_el.makeelement(qn("w:r"), {})
    t1 = r1.makeelement(qn("w:t"), {})
    t1.text = label + " "
    r1.append(t1)
    p_el.append(r1)

    # run2: STYLEREF 域（章节号）
    r2 = p_el.makeelement(qn("w:r"), {})
    p_el.append(r2)
    _add_field(r2, "STYLEREF 1 \\s")

    # run3: "-"
    r3 = p_el.makeelement(qn("w:r"), {})
    t3 = r3.makeelement(qn("w:t"), {})
    t3.text = "-"
    r3.append(t3)
    p_el.append(r3)

    # run4: SEQ 域（图号）
    r4 = p_el.makeelement(qn("w:r"), {})
    p_el.append(r4)
    _add_field(r4, f"SEQ {seq_name} \\* ARABIC \\s1")

    return p_el


def format_images(doc: Document, is_chinese: bool = True) -> dict:
    """批量图片排版（对齐 VBA 批量图片排版）。"""
    stats = {"total": 0, "processed": 0, "skipped": 0}

    # 遍历副本，避免插入题注后索引错乱
    shapes = list(doc.inline_shapes)
    for shape in shapes:
        stats["total"] += 1
        p_el = _shape_paragraph_el(shape)
        if p_el is None:
            stats["skipped"] += 1
            continue

        # 跳过已赋样式
        cur_name = _para_style_name(doc, p_el)
        if cur_name == STYLE_IMAGE:
            stats["skipped"] += 1
            continue

        w = shape.width or 0
        h = shape.height or 0
        if w > MIN_W_EMU and h > MIN_H_EMU:
            # 1. 段落居中
            _set_para_align_el(p_el, "center")
            # 2. 图片边框
            _set_image_border(shape)
            # 3. 赋样式
            _set_para_style_by_name(doc, p_el, STYLE_IMAGE)
            # 4. 插入题注到图片段落之后
            cap_p = _build_caption_paragraph(doc, is_chinese)
            p_el.addnext(cap_p)
            stats["processed"] += 1
        else:
            stats["skipped"] += 1
    return stats


# ========== 3. 清理标题多余编号 ==========

def _is_chinese_digit(ch: str) -> bool:
    return ch in "零一二三四五六七八九十百千两〇"


def _leading_number_len(s: str) -> int:
    """计算字符串开头"手动编号"占的字符数（含尾随分隔符）；不是编号返回 0。

    移植自 VBA 前导编号长度 函数。
    """
    if not s:
        return 0
    i = 0
    n = len(s)
    # 跳过前导空格/制表
    while i < n and s[i] in (" ", "\t", "　"):
        i += 1
    # 跳过"第"
    if i < n and s[i] == "第":
        i += 1
    # 数字部分
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
    # 章/节
    if i < n and s[i] in ("章", "节"):
        i += 1
    # 尾随分隔符
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
    """删除段落前 n 个字符（跨 run）。"""
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
    """清理标题多余编号（对齐 VBA 清理标题多余编号）。

    仅当标题段落已带自动编号（w:numPr）时，才删除文本开头的手动编号。
    """
    stats = {"cleaned": 0, "skipped": 0, "not_heading": 0}
    for para in doc.paragraphs:
        p_el = para._p
        name = _para_style_name(doc, p_el)
        # 仅处理 heading 1~9
        if not (name.startswith("heading ") and name[8:].isdigit()):
            stats["not_heading"] += 1
            continue
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            stats["skipped"] += 1
            continue
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            # 无自动编号，不清理
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
