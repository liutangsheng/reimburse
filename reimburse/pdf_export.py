from __future__ import annotations

import io
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

from reimburse.config import PAGE_H, PAGE_W, amount_to_chinese_units
from reimburse.ocr import _image_to_rgb_png


def load_signature_transparent(sig_path: Path) -> ImageReader:
    """Load signature image and make near-white background transparent."""
    img = Image.open(sig_path).convert("RGBA")
    raw = img.get_flattened_data() if hasattr(img, "get_flattened_data") else img.getdata()
    new_pixels = [
        (r, g, b, 0) if (r > 225 and g > 225 and b > 225) else (r, g, b, a)
        for (r, g, b, a) in raw
    ]
    img.putdata(new_pixels)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return ImageReader(buf)


def draw_circled_x(c: canvas.Canvas, cx: float, cy: float, radius: float = 4.2) -> None:
    """Draw vector ⓧ symbol so it never depends on font glyph support."""
    c.saveState()
    c.setLineWidth(0.7)
    c.circle(cx, cy, radius, stroke=1, fill=0)
    d = radius * 0.55
    c.line(cx - d, cy - d, cx + d, cy + d)
    c.line(cx - d, cy + d, cx + d, cy - d)
    c.restoreState()


def draw_centered_text(c: canvas.Canvas, x: float, y: float, text: str, font: str, size: float) -> None:
    c.setFont(font, size)
    c.drawCentredString(x, y - size * 0.35, text)


def draw_reimbursement_page(
    font_name: str,
    dept: str,
    fill_date: tuple[int, int, int],
    items: list[tuple[str, float]],
    remark: str,
    sig_path: Path | None,
    attachment_pages: int = 0,
) -> bytes:
    """Render Page 1 (费用报销单) as a vector PDF in memory."""
    total_amount = round(sum(amt for _, amt in items), 2)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    title_y = PAGE_H - 74
    c.setFont(font_name, 20)
    c.drawCentredString(PAGE_W / 2, title_y, "费 用 报 销 单")
    uw = 162
    ux0 = (PAGE_W - uw) / 2
    c.setLineWidth(1.4)
    c.line(ux0, title_y - 4, ux0 + uw, title_y - 4)
    c.setLineWidth(0.7)
    c.line(ux0, title_y - 7, ux0 + uw, title_y - 7)

    margin_x = 52
    table_w = PAGE_W - 2 * margin_x
    sub_y = title_y - 26
    c.setFont(font_name, 10)
    c.drawString(margin_x + 4, sub_y, "报销部门：")
    c.drawString(margin_x + 82, sub_y, dept)

    yr, mo, dy = fill_date
    c.drawCentredString(PAGE_W / 2 + 12, sub_y, f"{yr}  年  {mo}  月  {dy}  日填")

    right_x = margin_x + table_w - 115
    c.drawString(right_x, sub_y, "单据及附件共")
    c.setLineWidth(0.6)
    c.line(right_x + 62, sub_y - 2, right_x + 92, sub_y - 2)
    if attachment_pages > 0:
        c.drawCentredString(right_x + 77, sub_y, str(attachment_pages))
    c.drawString(right_x + 95, sub_y, "页")

    row_h = 23.0
    n_rows = 8
    table_top = sub_y - 6
    table_bottom = table_top - n_rows * row_h

    col_w = [218.0, 78.0, 22.0, 173.2]
    col_x = [margin_x]
    for w in col_w:
        col_x.append(col_x[-1] + w)

    c.setLineWidth(1.5)
    c.rect(margin_x, table_bottom, table_w, n_rows * row_h)

    c.setLineWidth(0.6)
    for r in range(1, n_rows):
        y = table_top - r * row_h
        if r in (2, 4, 5):
            c.line(col_x[0], y, col_x[2], y)
        else:
            c.line(col_x[0], y, col_x[4], y)

    c.line(col_x[1], table_top, col_x[1], table_bottom + row_h)
    c.line(col_x[2], table_top, col_x[2], table_bottom + row_h)
    c.line(col_x[3], table_top, col_x[3], table_bottom + 2 * row_h)

    r7_x1 = col_x[0] + 302
    r7_x2 = r7_x1 + 94
    c.line(r7_x1, table_bottom, r7_x1, table_bottom + row_h)
    c.line(r7_x2, table_bottom, r7_x2, table_bottom + row_h)

    mid_y = lambda r: table_top - (r + 0.5) * row_h

    draw_centered_text(c, (col_x[0] + col_x[1]) / 2, mid_y(0), "用      途", font_name, 10)
    draw_centered_text(c, (col_x[1] + col_x[2]) / 2, mid_y(0), "金额（元）", font_name, 10)

    bx = (col_x[2] + col_x[3]) / 2
    draw_centered_text(c, bx, table_top - 1.6 * row_h, "备", font_name, 10)
    draw_centered_text(c, bx, table_top - 2.4 * row_h, "注", font_name, 10)

    lx = (col_x[2] + col_x[3]) / 2
    for idx, ch in enumerate("领导审核"):
        draw_centered_text(c, lx, table_top - (3.4 + idx * 0.72) * row_h, ch, font_name, 9.5)

    for row_idx, (item_name, item_amt) in enumerate(items[:5], start=1):
        draw_centered_text(c, (col_x[0] + col_x[1]) / 2, mid_y(row_idx), item_name, font_name, 10)
        draw_centered_text(c, (col_x[1] + col_x[2]) / 2, mid_y(row_idx), f"{item_amt:.2f}", font_name, 10)

    if remark:
        rx = (col_x[3] + col_x[4]) / 2
        ry = table_top - 2.0 * row_h
        max_w = col_w[3] - 12
        font_sz = 10.0
        while font_sz > 7.5 and pdfmetrics.stringWidth(remark, font_name, font_sz) > max_w:
            font_sz -= 0.5
        if pdfmetrics.stringWidth(remark, font_name, font_sz) <= max_w:
            draw_centered_text(c, rx, ry, remark, font_name, font_sz)
        else:
            mid_idx = len(remark) // 2
            draw_centered_text(c, rx, ry + font_sz * 0.65, remark[:mid_idx], font_name, font_sz)
            draw_centered_text(c, rx, ry - font_sz * 0.65, remark[mid_idx:], font_name, font_sz)

    draw_centered_text(c, (col_x[0] + col_x[1]) / 2, mid_y(6), "合      计", font_name, 10)
    draw_centered_text(c, (col_x[1] + col_x[2]) / 2, mid_y(6), f"{total_amount:.2f}", font_name, 10)

    r7_y = mid_y(7) - 3.5
    c.setFont(font_name, 9.5)
    c.drawString(col_x[0] + 6, r7_y, "金额大写:")
    units = amount_to_chinese_units(total_amount)
    ux = col_x[0] + 54
    slot_w = 34.0
    for val, unit in units:
        c.setLineWidth(0.5)
        c.line(ux, r7_y - 2.5, ux + 20, r7_y - 2.5)
        if val == "ⓧ":
            draw_circled_x(c, ux + 10, r7_y + 4.0)
        elif val:
            c.drawCentredString(ux + 10, r7_y + 1.5, val)
        c.drawString(ux + 22, r7_y, unit)
        ux += slot_w

    c.drawString(r7_x1 + 6, r7_y, "原借款:")
    c.line(r7_x1 + 42, r7_y - 2, r7_x1 + 76, r7_y - 2)
    c.drawString(r7_x1 + 78, r7_y, "元")

    c.drawString(r7_x2 + 6, r7_y, "应退余款:")
    c.line(r7_x2 + 52, r7_y - 2, r7_x2 + 80, r7_y - 2)
    c.drawString(r7_x2 + 82, r7_y, "元")

    footer_y = table_bottom - 18
    c.setFont(font_name, 10)
    c.drawString(margin_x + 4, footer_y, "会计主管：")
    c.drawString(margin_x + 125, footer_y, "复核：")
    c.drawString(margin_x + 255, footer_y, "出纳：")
    reim_x = margin_x + 375
    c.drawString(reim_x, footer_y, "报销人：")

    if sig_path and sig_path.exists():
        sig_reader = load_signature_transparent(sig_path)
        iw, ih = sig_reader.getSize()
        target_w = 56.0
        target_h = target_w * (ih / iw)
        c.drawImage(
            sig_reader,
            reim_x + 40,
            footer_y - target_h * 0.55,
            width=target_w,
            height=target_h,
            mask="auto",
        )

    c.showPage()
    c.save()
    return buf.getvalue()


def draw_order_pages(order_images: list[Path]) -> list[bytes]:
    """Render order screenshots onto landscape pages, up to 3 per page."""
    if not order_images:
        return []

    pages = []
    margin_x = 42.0
    margin_y = 42.0
    gap = 18.0
    cols = 3
    slot_w = (PAGE_W - 2 * margin_x - (cols - 1) * gap) / cols
    slot_h = PAGE_H - 2 * margin_y

    for i in range(0, len(order_images), cols):
        chunk = order_images[i : i + cols]
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
        for idx, img_path in enumerate(chunk):
            with Image.open(img_path) as im:
                iw, ih = im.size
            scale = min(slot_w / iw, slot_h / ih)
            draw_w = iw * scale
            draw_h = ih * scale
            x = margin_x + idx * (slot_w + gap) + (slot_w - draw_w) / 2
            y = margin_y + (slot_h - draw_h) / 2
            c.drawImage(str(img_path), x, y, width=draw_w, height=draw_h)
        c.showPage()
        c.save()
        pages.append(buf.getvalue())
    return pages


def draw_invoice_image_page(img_path: Path) -> bytes:
    """Render an invoice image onto a single PDF page (fitting A5 landscape with margin)."""
    img_buf, iw, ih = _image_to_rgb_png(img_path)
    margin = 18.0
    avail_w = PAGE_W - 2 * margin
    avail_h = PAGE_H - 2 * margin
    scale = min(avail_w / iw, avail_h / ih)
    draw_w = iw * scale
    draw_h = ih * scale
    x = (PAGE_W - draw_w) / 2
    y = (PAGE_H - draw_h) / 2

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    c.drawImage(ImageReader(img_buf), x, y, width=draw_w, height=draw_h)
    c.showPage()
    c.save()
    return buf.getvalue()


def _append_invoice_to_writer(writer: PdfWriter, inv_path: Path) -> None:
    """Append either a PDF invoice or an image invoice to the output PdfWriter."""
    if inv_path.suffix.lower() == ".pdf":
        for page in PdfReader(str(inv_path)).pages:
            writer.add_page(page)
    else:
        page_bytes = draw_invoice_image_page(inv_path)
        for page in PdfReader(io.BytesIO(page_bytes)).pages:
            writer.add_page(page)
