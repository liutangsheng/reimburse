from __future__ import annotations

import io
from pathlib import Path

from PIL import Image

from reimburse.config import PAGE_H, PAGE_W, amount_to_chinese_units
from reimburse.ocr import _image_to_rgb_png, _render_pdf_pages_to_pngs


def _set_cell_font(run, size_pt: float = 10.0, bold: bool = False, name: str = "宋体") -> None:
    """Apply consistent Chinese font styling to a python-docx Run."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    run.font.name = name
    run.font.size = Pt(size_pt)
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), name)


def _format_docx_cell(
    cell,
    text: str = "",
    size_pt: float = 10.0,
    align: str = "center",
) -> None:
    """Set centered/aligned text and vertical centering on a python-docx table cell."""
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    if text:
        run = p.add_run(text)
        _set_cell_font(run, size_pt=size_pt)


def _set_table_borders(table) -> None:
    """Apply thick outer border and thin inner borders to the reimbursement Word table."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (
        ("top", "12"),
        ("left", "12"),
        ("bottom", "12"),
        ("right", "12"),
        ("insideH", "4"),
        ("insideV", "4"),
    ):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tbl_pr.append(borders)


def _add_docx_fullpage_image(doc, img_buf: io.BytesIO, max_w_pt: float, max_h_pt: float) -> None:
    """Append a page break and a centered image scaled to fit within max_w_pt x max_h_pt."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    doc.add_page_break()
    img_buf.seek(0)
    with Image.open(img_buf) as im:
        iw, ih = im.size
    scale = min(max_w_pt / iw, max_h_pt / ih)
    draw_w = iw * scale
    img_buf.seek(0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(img_buf, width=Pt(draw_w))


def export_reimbursement_docx(
    output_docx: Path,
    dept: str,
    fill_date: tuple[int, int, int],
    items: list[tuple[str, float]],
    remark: str,
    sig_path: Path | None,
    order_images: list[Path] | None = None,
    order_pdfs: list[Path] | None = None,
    invoice_files: list[Path] | None = None,
    attachment_pages: int = 0,
) -> None:
    """Export the full reimbursement package (Form + Orders + Invoices) as an editable Word (.docx)."""
    from docx import Document
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
    from docx.shared import Pt

    total_amount = round(sum(amt for _, amt in items), 2)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Pt(PAGE_W)
    sec.page_height = Pt(PAGE_H)
    sec.left_margin = Pt(52)
    sec.right_margin = Pt(52)
    sec.top_margin = Pt(36)
    sec.bottom_margin = Pt(30)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("费 用 报 销 单")
    r_title.underline = WD_UNDERLINE.DOUBLE
    _set_cell_font(r_title, size_pt=20, bold=False)

    yr, mo, dy = fill_date
    pages_str = f" {attachment_pages} " if attachment_pages > 0 else " ____ "
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(4)
    r_sub = p_sub.add_run(
        f"报销部门：  {dept:<14}    {yr} 年 {mo} 月 {dy} 日填        单据及附件共{pages_str}页"
    )
    _set_cell_font(r_sub, size_pt=10)

    table = doc.add_table(rows=8, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_borders(table)

    col_widths = [Pt(218), Pt(78), Pt(22), Pt(173.2)]
    for row in table.rows:
        row.height = Pt(23)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for idx, w in enumerate(col_widths):
            row.cells[idx].width = w

    _format_docx_cell(table.cell(0, 0), "用      途", 10)
    _format_docx_cell(table.cell(0, 1), "金额（元）", 10)

    c_bz_lbl = table.cell(1, 2).merge(table.cell(2, 2))
    _format_docx_cell(c_bz_lbl, "备\n注", 10)
    c_bz_val = table.cell(1, 3).merge(table.cell(2, 3))
    _format_docx_cell(c_bz_val, remark, 10)

    c_ld_lbl = table.cell(3, 2).merge(table.cell(5, 2))
    _format_docx_cell(c_ld_lbl, "领\n导\n审\n核", 9.5)
    table.cell(3, 3).merge(table.cell(5, 3))

    for row_idx, (item_name, item_amt) in enumerate(items[:5], start=1):
        _format_docx_cell(table.cell(row_idx, 0), item_name, 10)
        _format_docx_cell(table.cell(row_idx, 1), f"{item_amt:.2f}", 10)

    _format_docx_cell(table.cell(6, 0), "合      计", 10)
    _format_docx_cell(table.cell(6, 1), f"{total_amount:.2f}", 10)
    table.cell(6, 2).merge(table.cell(6, 3))

    c_r7 = table.cell(7, 0).merge(table.cell(7, 3))
    units = amount_to_chinese_units(total_amount)
    upper_parts = []
    for val, unit in units:
        mark = "⊗" if val == "ⓧ" else (val or "  ")
        upper_parts.append(f" {mark} {unit}")
    r7_text = f"金额大写:{''.join(upper_parts)}   |   原借款: ____ 元   |   应退余款: ____ 元"
    _format_docx_cell(c_r7, r7_text, 9.0, align="left")

    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(8)
    p_foot.paragraph_format.space_after = Pt(0)
    r_foot = p_foot.add_run("会计主管：            复核：            出纳：            报销人：")
    _set_cell_font(r_foot, size_pt=10)

    if sig_path and sig_path.exists():
        sig_img = Image.open(sig_path).convert("RGBA")
        raw = sig_img.get_flattened_data() if hasattr(sig_img, "get_flattened_data") else sig_img.getdata()
        sig_img.putdata([
            (r, g, b, 0) if (r > 225 and g > 225 and b > 225) else (r, g, b, a)
            for (r, g, b, a) in raw
        ])
        sig_buf = io.BytesIO()
        sig_img.save(sig_buf, format="PNG")
        sig_buf.seek(0)
        p_foot.add_run().add_picture(sig_buf, width=Pt(56))

    avail_w = PAGE_W - 104.0
    avail_h = PAGE_H - 72.0
    if order_images:
        slot_w = (avail_w - 20.0) / 3.0
        slot_h = avail_h - 10.0
        for i in range(0, len(order_images), 3):
            chunk = order_images[i : i + 3]
            doc.add_page_break()
            tbl = doc.add_table(rows=1, cols=3)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = False
            for idx, img_path in enumerate(chunk):
                cell = tbl.cell(0, idx)
                cell.width = Pt(slot_w)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                buf, iw, ih = _image_to_rgb_png(img_path)
                scale = min((slot_w - 6.0) / iw, slot_h / ih)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.add_run().add_picture(buf, width=Pt(iw * scale))

    for pdf_path in order_pdfs or []:
        for png_buf in _render_pdf_pages_to_pngs(pdf_path):
            _add_docx_fullpage_image(doc, png_buf, avail_w, avail_h)

    for inv_path in invoice_files or []:
        if inv_path.suffix.lower() == ".pdf":
            for png_buf in _render_pdf_pages_to_pngs(inv_path):
                _add_docx_fullpage_image(doc, png_buf, avail_w, avail_h)
        else:
            buf, _, _ = _image_to_rgb_png(inv_path)
            _add_docx_fullpage_image(doc, buf, avail_w, avail_h)

    doc.save(str(output_docx))
