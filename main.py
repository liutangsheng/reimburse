from __future__ import annotations

import argparse
import contextlib
import io
import os
import re
import shutil
import subprocess
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

PAGE_W, PAGE_H = 595.2, 419.52  # A5 landscape (matches 费用报销单.docx & reference PDF)
CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
MIN_ORDER_TIME = (17, 30)    # 餐饮订单 17:30 以后下单才有效
MIN_TRAFFIC_TIME = (23, 0)   # 交通行程 23:00 (晚上11:00) 以后才有效
MAX_ORDER_AGE_DAYS = 14      # 下单/出行时间与当前时间不能超过两周
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
INVOICE_EXTS = {".pdf"} | IMAGE_EXTS

_TESS_HAS_CHI_SIM: bool | None = None


def register_chinese_font() -> str:
    """Register a system Chinese TrueType font for ReportLab (macOS & Windows)."""
    candidates = [
        ("Songti", "/System/Library/Fonts/Supplemental/Songti.ttc", 0),
        ("STHeiti", "/System/Library/Fonts/STHeiti Medium.ttc", 0),
        ("SimSun", r"C:\Windows\Fonts\simsun.ttc", 0),
        ("SimSunTTF", r"C:\Windows\Fonts\simsun.ttf", None),
        ("MSYaHei", r"C:\Windows\Fonts\msyh.ttc", 0),
    ]
    for name, path, idx in candidates:
        if Path(path).exists():
            kwargs = {"subfontIndex": idx} if idx is not None else {}
            pdfmetrics.registerFont(TTFont(name, path, **kwargs))
            return name
    return "Helvetica"


def _find_tesseract() -> str:
    """Locate tesseract binary even when launched from Finder/Explorer without PATH."""
    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in (
        "/opt/homebrew/bin/tesseract",
        "/usr/local/bin/tesseract",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return "tesseract"


def _has_tesseract_chi_sim() -> bool:
    """Check whether tesseract has the chi_sim language pack installed."""
    global _TESS_HAS_CHI_SIM
    if _TESS_HAS_CHI_SIM is None:
        try:
            res = subprocess.run(
                [_find_tesseract(), "--list-langs"],
                capture_output=True,
                text=True,
                check=False,
            )
            langs = {line.strip() for line in (res.stdout or "").splitlines()}
            _TESS_HAS_CHI_SIM = "chi_sim" in langs
        except OSError:
            _TESS_HAS_CHI_SIM = False
    return _TESS_HAS_CHI_SIM


def _to_grayscale_on_white(im: Image.Image) -> Image.Image:
    """Convert a PIL image to grayscale, compositing transparent pixels over white first."""
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        rgba = im.convert("RGBA")
        bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        return Image.alpha_composite(bg, rgba).convert("L")
    return im.convert("L")


def _ocr_image(
    img: Image.Image,
    psm: int | None = None,
    work_dir: Path | None = None,
    lang: str = "eng",
) -> str:
    """Run tesseract OCR on a PIL Image using a temporary file in the current directory."""
    base_dir = Path.cwd()
    if not os.access(base_dir, os.W_OK) and work_dir is not None:
        base_dir = work_dir
    tmp_path = base_dir / f"_ocr_tmp_{uuid.uuid4().hex}.png"
    cmd = [_find_tesseract(), str(tmp_path), "stdout", "-l", lang]
    if psm is not None:
        cmd.extend(["--psm", str(psm)])
    try:
        img.save(tmp_path)
        res = subprocess.run(cmd, capture_output=True, text=True, check=False)
        return res.stdout or ""
    finally:
        with contextlib.suppress(OSError):
            tmp_path.unlink(missing_ok=True)


def parse_order_image(img_path: Path) -> dict:
    """Extract order datetime and amount from an order screenshot via Tesseract OCR."""
    work_dir = img_path.parent
    with Image.open(img_path) as im:
        gray = _to_grayscale_on_white(im)
        full_text = _ocr_image(gray, work_dir=work_dir)

        # Upscale bottom 28% (Meituan 合计 ¥xx.x) and middle 35% (Ele.me 实付¥xx.x) for accurate decimal OCR
        bottom = gray.crop((0, int(gray.height * 0.72), gray.width, gray.height))
        bottom_3x = bottom.resize((bottom.width * 3, bottom.height * 3), Image.Resampling.LANCZOS)
        bottom_text = _ocr_image(bottom_3x, psm=6, work_dir=work_dir)

        mid = gray.crop((0, int(gray.height * 0.35), gray.width, int(gray.height * 0.70)))
        mid_2x = mid.resize((mid.width * 2, mid.height * 2), Image.Resampling.LANCZOS)
        mid_text = _ocr_image(mid_2x, psm=6, work_dir=work_dir)

    # Extract full datetime (e.g. 2026-11-20 18:28:52)
    dt_m = re.search(r"(20\d{2}-\d{2}-\d{2})\s+(\d{2}:\d{2})(?::(\d{2}))?", full_text)
    order_dt = None
    if dt_m:
        sec = dt_m.group(3) or "00"
        try:
            order_dt = datetime.strptime(f"{dt_m.group(1)} {dt_m.group(2)}:{sec}", "%Y-%m-%d %H:%M:%S")
        except ValueError:
            order_dt = None
    date_only = dt_m.group(1) if dt_m else None
    if not date_only:
        d_m = re.search(r"(20\d{2}-\d{2}-\d{2})", full_text)
        date_only = d_m.group(1) if d_m else None

    # Extract amount: prefer ¥/￥ prefixed decimals in bottom_text, then mid_text, then full_text
    amount = None
    for txt in (bottom_text, mid_text, full_text):
        yen_nums = re.findall(r"[¥￥]\s*(\d{1,4}\.\d{1,2})", txt)
        if yen_nums:
            amount = float(yen_nums[-1])
            break
    if amount is None:
        for txt in (bottom_text, mid_text, full_text):
            decimals = re.findall(r"\b(\d{1,4}\.\d{1,2})\b", txt)
            if decimals:
                amount = float(decimals[-1])
                break

    return {
        "path": img_path,
        "date": date_only,
        "datetime": order_dt,
        "amount": amount,
    }


def _extract_delivery_fee(full_text: str) -> float:
    """Detect and sum any 配送费 / 配送服务 line items (tax-inclusive) in invoice text."""
    fee_total = 0.0
    for line in full_text.splitlines():
        if "配送费" not in line and "配送服务" not in line:
            continue
        rate_m = re.search(r"(\d+(?:\.\d+)?)%", line)
        rate = float(rate_m.group(1)) / 100.0 if rate_m else None
        # Extract numbers not followed by '%'
        nums = [float(x) for x in re.findall(r"(?<!\d)(\d+\.\d{1,2})(?!\d|%)", line)]
        if not nums:
            continue
        if len(nums) >= 2 and rate is not None and abs(nums[-1] - nums[-2] * rate) <= 0.05:
            fee_total += nums[-2] + nums[-1]
        else:
            fee_total += nums[-1]
    return round(fee_total, 2)


def _find_yen_amount(text: str) -> float | None:
    """Match either '¥56.20', '56.20¥', or OCR variants like '¥43, 00' / '¥42. 57'."""
    m = re.search(
        r"[¥￥]\s*(-?\d+(?:\s*[.,]\s*\d{1,2})?)|(-?\d+(?:\s*[.,]\s*\d{1,2})?)\s*[¥￥]",
        text,
    )
    if not m:
        return None
    raw = m.group(1) or m.group(2)
    norm = re.sub(r"\s*[.,]\s*", ".", raw)
    try:
        return float(norm)
    except ValueError:
        return None


def _parse_invoice_text(
    full_text: str,
    inv_path: Path,
    pages: int = 1,
    is_image: bool = False,
    has_chi_sim: bool = True,
    default_purpose: str = "餐费",
) -> dict:
    """Parse invoice fields from extracted PDF text or image OCR text."""
    # Prefer the ¥ amount on the line with Chinese uppercase total (e.g. 肆拾叁圆整 ¥43.00 or 56.20¥)
    raw_amount = None
    for line in full_text.splitlines():
        if any(k in line for k in ("圆整", "角", "分", "价税合计", "小写")):
            amt = _find_yen_amount(line)
            if amt is not None:
                raw_amount = amt
                break

    if raw_amount is None:
        all_amounts = []
        for m in re.finditer(
            r"[¥￥]\s*(\d+(?:\s*[.,]\s*\d{1,2})?)|(\d+(?:\s*[.,]\s*\d{1,2})?)\s*[¥￥]",
            full_text,
        ):
            raw = m.group(1) or m.group(2)
            norm = re.sub(r"\s*[.,]\s*", ".", raw)
            with contextlib.suppress(ValueError):
                all_amounts.append(float(norm))
        if all_amounts:
            raw_amount = max(all_amounts)

    if raw_amount is None and is_image:
        decimals = [float(x) for x in re.findall(r"\b(\d{1,5}\.\d{2})\b", full_text)]
        if decimals:
            raw_amount = max(decimals)

    delivery_fee = _extract_delivery_fee(full_text)
    effective_amount = (
        round(max(0.0, raw_amount - delivery_fee), 2)
        if raw_amount is not None
        else None
    )

    inv_date = None
    date_m = (
        re.search(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", full_text)
        or re.search(r"(20[2-3]\d)[-./](1[0-2]|0?[1-9])[-./]([12]\d|3[01]|0?[1-9])", full_text)
        or re.search(
            r"(20[2-3]\d)[457]{1,2}(1[0-2]|0?[1-9])[59A)\s]+([12]\d|3[01]|0?[1-9])[8B1H](?![0-9A-Za-z])",
            full_text,
        )
        or re.search(
            r"(?<!\d)(20[2-3]\d)[-.]?(0[1-9]|1[0-2])[-.]?(0[1-9]|[12]\d|3[01])",
            inv_path.name,
        )
    )
    if date_m:
        inv_date = f"{date_m.group(1)}-{int(date_m.group(2)):02d}-{int(date_m.group(3)):02d}"

    is_guangzhou = "广州" in full_text or "广州" in inv_path.name
    if not is_guangzhou and is_image:
        # Check Unified Social Credit Code prefix for Guangzhou (914401 / 924401)
        gz_codes = [
            re.sub(r"\s+", "", m.group(0))
            for m in re.finditer(r"9[12]4401\d{2}\s*[0-9A-Z]{10}", full_text)
        ]
        seller_gz = [c for c in gz_codes if c != "914401055622643150"]
        if seller_gz or len(gz_codes) >= 2:
            is_guangzhou = True

    combined = f"{full_text}\n{inv_path.name}"
    is_catering = "餐饮" in combined
    is_traffic = any(k in combined for k in ("交通", "客运", "出行")) or "didi" in combined.lower()
    is_hotel = "住宿" in combined

    if not is_catering and is_image and not has_chi_sim and default_purpose == "餐费":
        if not is_traffic and not is_hotel and raw_amount is not None:
            is_catering = True

    purpose = default_purpose
    if is_traffic:
        purpose = "交通费"
    elif is_hotel:
        purpose = "住宿费"
    elif is_catering:
        purpose = "餐费"

    return {
        "path": inv_path,
        "raw_amount": raw_amount,
        "delivery_fee": delivery_fee,
        "amount": effective_amount,
        "date": inv_date,
        "is_guangzhou": is_guangzhou,
        "is_catering": is_catering,
        "purpose": purpose,
        "pages": pages,
    }


def _ocr_invoice_pil(im: Image.Image, work_dir: Path | None = None) -> tuple[str, bool]:
    """Run OCR on a PIL image of an invoice and return (combined_text, has_chi_sim)."""
    gray = _to_grayscale_on_white(im)
    if gray.width < 1600:
        scale = max(2, round(1800 / max(1, gray.width)))
        gray = gray.resize((gray.width * scale, gray.height * scale), Image.Resampling.LANCZOS)

    has_chi = _has_tesseract_chi_sim()
    texts = []
    if has_chi:
        texts.append(_ocr_image(gray, work_dir=work_dir, lang="chi_sim+eng"))
    texts.append(_ocr_image(gray, work_dir=work_dir, lang="eng"))
    texts.append(_ocr_image(gray, psm=6, work_dir=work_dir, lang="eng"))
    return "\n".join(texts), has_chi


def parse_invoice_image(img_path: Path, default_purpose: str = "餐费") -> dict:
    """Extract total amount, delivery fee, seller, date, and category from an invoice image."""
    with Image.open(img_path) as im:
        full_text, has_chi = _ocr_invoice_pil(im, work_dir=img_path.parent)
    return _parse_invoice_text(
        full_text,
        img_path,
        pages=1,
        is_image=True,
        has_chi_sim=has_chi,
        default_purpose=default_purpose,
    )


def parse_invoice_pdf(pdf_path: Path, default_purpose: str = "餐费") -> dict:
    """Extract total amount, delivery fee, seller, date, and category from an invoice PDF."""
    reader = PdfReader(str(pdf_path))
    full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

    if not full_text.strip():
        # Fallback for scanned image-only PDFs
        ocr_parts = []
        has_chi = _has_tesseract_chi_sim()
        for page in reader.pages:
            for img_file in getattr(page, "images", []):
                with Image.open(io.BytesIO(img_file.data)) as im:
                    t, has_chi = _ocr_invoice_pil(im, work_dir=pdf_path.parent)
                    ocr_parts.append(t)
        if ocr_parts:
            return _parse_invoice_text(
                "\n".join(ocr_parts),
                pdf_path,
                pages=len(reader.pages),
                is_image=True,
                has_chi_sim=has_chi,
                default_purpose=default_purpose,
            )

    return _parse_invoice_text(
        full_text,
        pdf_path,
        pages=len(reader.pages),
        is_image=False,
        has_chi_sim=True,
        default_purpose=default_purpose,
    )


def parse_invoice_file(inv_path: Path, default_purpose: str = "餐费") -> dict:
    """Extract invoice information from either a PDF or an image invoice file."""
    if inv_path.suffix.lower() == ".pdf":
        return parse_invoice_pdf(inv_path, default_purpose=default_purpose)
    return parse_invoice_image(inv_path, default_purpose=default_purpose)


def parse_traffic_trip_pdf(pdf_path: Path) -> dict:
    """Extract trip datetimes and amounts from a ride-hailing trip sheet PDF (e.g. 滴滴出行行程单)."""
    reader = PdfReader(str(pdf_path))
    full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

    year_m = re.search(r"(?:行程起止日期|申请日期)[：:]\s*(20\d{2})-", full_text)
    year = int(year_m.group(1)) if year_m else datetime.now().year

    total_m = re.search(r"合计\s*(\d+(?:\.\d{1,2})?)\s*元", full_text)
    declared_total = float(total_m.group(1)) if total_m else None

    trips = []
    for line in full_text.splitlines():
        m = re.search(r"(\d{2}-\d{2})\s+(\d{2}:\d{2}).*?(\d+\.\d{2})\s*$", line.strip())
        if m:
            mm_dd, hh_mm, amt_s = m.group(1), m.group(2), m.group(3)
            try:
                dt = datetime.strptime(f"{year}-{mm_dd} {hh_mm}:00", "%Y-%m-%d %H:%M:%S")
            except ValueError:
                dt = None
            trips.append({
                "datetime": dt,
                "date": f"{year}-{mm_dd}",
                "amount": float(amt_s),
            })

    return {
        "path": pdf_path,
        "trips": trips,
        "declared_total": declared_total,
        "pages": len(reader.pages),
    }


def classify_traffic_pdfs(traffic_dir: Path) -> tuple[list[Path], list[Path]]:
    """Split files in traffic/ into (trip_pdfs, invoice_files) supporting PDF and image invoices."""
    trip_pdfs: list[Path] = []
    invoice_files: list[Path] = []
    if not traffic_dir.is_dir():
        return trip_pdfs, invoice_files

    for p in sorted(traffic_dir.iterdir()):
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if ext == ".pdf":
            text = "\n".join(page.extract_text() or "" for page in PdfReader(str(p)).pages)
            if "电子发票" in text or "发票号码" in text or "发票" in p.name:
                invoice_files.append(p)
            else:
                trip_pdfs.append(p)
        elif ext in IMAGE_EXTS:
            invoice_files.append(p)
    return trip_pdfs, invoice_files


def format_date_range(dates: list[str]) -> str:
    """Format sorted YYYY-MM-DD dates into 'MM.DD-MM.DD 号加班' or 'MM.DD 号加班'."""
    if not dates:
        return ""
    sorted_dates = sorted(set(dates))
    fmt = lambda d: f"{int(d[5:7])}.{int(d[8:10])}"
    if len(sorted_dates) == 1:
        return f"{fmt(sorted_dates[0])} 号加班"
    return f"{fmt(sorted_dates[0])}-{fmt(sorted_dates[-1])} 号加班"


def amount_to_chinese_units(amount: float) -> list[tuple[str, str]]:
    """Return [(digit_or_mark, unit_label), ...] for 万 仟 佰 拾 元 角 分."""
    cents = round(amount * 100)
    units = ["万", "仟", "佰", "拾", "元", "角", "分"]
    divisors = [1000000, 100000, 10000, 1000, 100, 10, 1]
    digits = []
    rem = cents
    for div in divisors:
        d, rem = divmod(rem, div)
        digits.append(d)

    # Find highest non-zero digit index (at least 元 at index 4)
    first_nz = next((i for i, d in enumerate(digits) if d > 0), 4)
    result = []
    for i, (d, unit) in enumerate(zip(digits, units)):
        if i < first_nz - 1:
            val = ""
        elif i == first_nz - 1:
            val = "ⓧ"
        else:
            val = CN_DIGITS[d]
        result.append((val, unit))
    return result


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


def draw_centered_text(c: canvas.Canvas, x: float, y: float, text: str, font: str, size: float):
    c.setFont(font, size)
    c.drawCentredString(x, y - size * 0.35, text)


def draw_reimbursement_page(
    font_name: str,
    dept: str,
    fill_date: tuple[int, int, int],
    items: list[tuple[str, float]],
    remark: str,
    sig_path: Path | None,
) -> bytes:
    """Render Page 1 (费用报销单) as a vector PDF in memory."""
    total_amount = round(sum(amt for _, amt in items), 2)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    # Title
    title_y = PAGE_H - 74
    c.setFont(font_name, 20)
    c.drawCentredString(PAGE_W / 2, title_y, "费 用 报 销 单")
    # Double underline under title
    uw = 162
    ux0 = (PAGE_W - uw) / 2
    c.setLineWidth(1.4)
    c.line(ux0, title_y - 4, ux0 + uw, title_y - 4)
    c.setLineWidth(0.7)
    c.line(ux0, title_y - 7, ux0 + uw, title_y - 7)

    # Sub-header line
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
    c.drawString(right_x + 95, sub_y, "页")

    # Table geometry (8 rows, matching 费用报销单.docx)
    row_h = 23.0
    n_rows = 8
    table_top = sub_y - 6
    table_bottom = table_top - n_rows * row_h

    # Column widths summing to table_w (491.2 pt)
    col_w = [218.0, 78.0, 22.0, 173.2]
    col_x = [margin_x]
    for w in col_w:
        col_x.append(col_x[-1] + w)

    # Outer thick border
    c.setLineWidth(1.5)
    c.rect(margin_x, table_bottom, table_w, n_rows * row_h)

    # Inner horizontal lines
    c.setLineWidth(0.6)
    for r in range(1, n_rows):
        y = table_top - r * row_h
        if r in (2, 4, 5):
            # Rows 1-2 (备注) and 3-5 (领导审核) are merged on columns 2 & 3
            c.line(col_x[0], y, col_x[2], y)
        else:
            c.line(col_x[0], y, col_x[4], y)

    # Inner vertical lines
    # Col 1 separator (用途 | 金额) for rows 0..6
    c.line(col_x[1], table_top, col_x[1], table_bottom + row_h)
    # Col 2 separator (金额 | 备注/领导审核) for rows 0..6
    c.line(col_x[2], table_top, col_x[2], table_bottom + row_h)
    # Col 3 separator (备注/领导审核 label | content) for rows 0..5
    c.line(col_x[3], table_top, col_x[3], table_bottom + 2 * row_h)

    # Row 7 bottom split: 金额大写 | 原借款 | 应退余款
    r7_x1 = col_x[0] + 302
    r7_x2 = r7_x1 + 94
    c.line(r7_x1, table_bottom, r7_x1, table_bottom + row_h)
    c.line(r7_x2, table_bottom, r7_x2, table_bottom + row_h)

    # Fill table text
    mid_y = lambda r: table_top - (r + 0.5) * row_h

    draw_centered_text(c, (col_x[0] + col_x[1]) / 2, mid_y(0), "用      途", font_name, 10)
    draw_centered_text(c, (col_x[1] + col_x[2]) / 2, mid_y(0), "金额（元）", font_name, 10)

    # 备注 (vertical across rows 0..2)
    bx = (col_x[2] + col_x[3]) / 2
    draw_centered_text(c, bx, table_top - 1.0 * row_h, "备", font_name, 10)
    draw_centered_text(c, bx, table_top - 2.0 * row_h, "注", font_name, 10)

    # 领导审核 (vertical across rows 3..5)
    lx = (col_x[2] + col_x[3]) / 2
    for idx, ch in enumerate("领导审核"):
        draw_centered_text(c, lx, table_top - (3.4 + idx * 0.72) * row_h, ch, font_name, 9.5)

    # Rows 1..5: fill items (up to 5 rows)
    for row_idx, (item_name, item_amt) in enumerate(items[:5], start=1):
        draw_centered_text(c, (col_x[0] + col_x[1]) / 2, mid_y(row_idx), item_name, font_name, 10)
        draw_centered_text(c, (col_x[1] + col_x[2]) / 2, mid_y(row_idx), f"{item_amt:.2f}", font_name, 10)

    # Remark centered in rows 0..2 col 3
    if remark:
        draw_centered_text(c, (col_x[3] + col_x[4]) / 2, table_top - 1.5 * row_h, remark, font_name, 10)

    # Row 6: 合计
    draw_centered_text(c, (col_x[0] + col_x[1]) / 2, mid_y(6), "合      计", font_name, 10)
    draw_centered_text(c, (col_x[1] + col_x[2]) / 2, mid_y(6), f"{total_amount:.2f}", font_name, 10)

    # Row 7: 金额大写
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

    # 原借款 / 应退余款
    c.drawString(r7_x1 + 6, r7_y, "原借款:")
    c.line(r7_x1 + 42, r7_y - 2, r7_x1 + 76, r7_y - 2)
    c.drawString(r7_x1 + 78, r7_y, "元")

    c.drawString(r7_x2 + 6, r7_y, "应退余款:")
    c.line(r7_x2 + 52, r7_y - 2, r7_x2 + 80, r7_y - 2)
    c.drawString(r7_x2 + 82, r7_y, "元")

    # Footer
    footer_y = table_bottom - 18
    c.setFont(font_name, 10)
    c.drawString(margin_x + 4, footer_y, "会计主管：")
    c.drawString(margin_x + 125, footer_y, "复核：")
    c.drawString(margin_x + 255, footer_y, "出纳：")
    reim_x = margin_x + 375
    c.drawString(reim_x, footer_y, "报销人：")

    # Signature image next to 报销人：
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
    with Image.open(img_path) as raw:
        if raw.mode in ("RGBA", "LA") or (raw.mode == "P" and "transparency" in raw.info):
            rgba = raw.convert("RGBA")
            bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
            im = Image.alpha_composite(bg, rgba).convert("RGB")
        else:
            im = raw.convert("RGB")
        iw, ih = im.size
        img_buf = io.BytesIO()
        im.save(img_buf, format="PNG")
        img_buf.seek(0)

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


def parse_dir_date(dir_name: str) -> tuple[int, int, int]:
    m = re.search(r"(20\d{2})\.(\d{1,2})\.(\d{1,2})", dir_name)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    now = datetime.now()
    return now.year, now.month, now.day


def _is_after_traffic_cutoff(dt: datetime) -> bool:
    """Valid if >= 23:00 (or early morning < 05:00)."""
    return (dt.hour, dt.minute) >= MIN_TRAFFIC_TIME or dt.hour < 5


def check_compliance(
    orders: list[dict],
    invoices: list[dict],
    traffic_trips: list[dict],
    traffic_invoices: list[dict],
    now: datetime | None = None,
) -> tuple[list[tuple[str, float]], list[str], list[str]]:
    """Run all compliance checks, log results, and return (items, valid_dates, issues)."""
    now = now or datetime.now()
    issues: list[str] = []
    items: list[tuple[str, float]] = []
    valid_dates: list[str] = []

    print("\n[3/4] 正在核对合规性...")
    print(f"  当前参考时间: {now.strftime('%Y-%m-%d %H:%M:%S')}")

    # --- 1. 餐饮订单 & 餐饮发票 ---
    valid_order_amounts: list[float] = []
    for idx, o in enumerate(orders, 1):
        name = o["path"].name
        dt = o["datetime"]
        amt = o["amount"]
        if dt is None:
            msg = f"餐饮订单 {idx} ({name}): 未能识别完整下单时间，无法校验 17:30 与两周限制"
            issues.append(msg)
            print(f"  [警告] {msg}")
            if o["date"]:
                valid_dates.append(o["date"])
            if amt is not None:
                valid_order_amounts.append(amt)
            continue

        after_1730 = (dt.hour, dt.minute) >= MIN_ORDER_TIME
        age = now - dt
        within_two_weeks = timedelta(0) <= age <= timedelta(days=MAX_ORDER_AGE_DAYS)
        age_days = age.total_seconds() / 86400

        status_parts = []
        if after_1730:
            status_parts.append("17:30后下单✓")
        else:
            status_parts.append("早于17:30✗")
            issues.append(f"餐饮订单 {idx} ({name}) 下单时间 {dt.strftime('%H:%M')} 早于 17:30，无效")

        if within_two_weeks:
            status_parts.append(f"距今{age_days:.1f}天(<=14天)✓")
        else:
            status_parts.append(f"距今{age_days:.1f}天(超14天)✗")
            issues.append(f"餐饮订单 {idx} ({name}) 下单时间 {dt.strftime('%Y-%m-%d')} 距今超过两周({age_days:.1f}天)")

        is_valid = after_1730 and within_two_weeks
        tag = "通过" if is_valid else "不合规"
        amt_str = f"¥{amt:.2f}" if amt is not None else "未识别"
        print(f"  [{tag}] 餐饮订单 {idx} ({name}): 时间 {dt.strftime('%Y-%m-%d %H:%M:%S')} | 金额 {amt_str} | {', '.join(status_parts)}")

        if is_valid:
            valid_dates.append(o["date"])
            if amt is not None:
                valid_order_amounts.append(amt)

    valid_invoice_amounts: list[float] = []
    for idx, inv in enumerate(invoices, 1):
        name = inv["path"].name
        gz_ok = inv["is_guangzhou"]
        cat_ok = inv["is_catering"]
        if not (gz_ok and cat_ok):
            issues.append(f"餐饮发票 {idx} ({name}) 不是广州餐饮票 (广州:{gz_ok}, 餐饮:{cat_ok})")
        tag = "通过" if (gz_ok and cat_ok) else "不合规"
        fee_note = (
            f"扣除配送费 ¥{inv['delivery_fee']:.2f}"
            if inv["delivery_fee"] > 0
            else "无配送费"
        )
        raw_str = f"¥{inv['raw_amount']:.2f}" if inv["raw_amount"] is not None else "未知"
        eff_str = f"¥{inv['amount']:.2f}" if inv["amount"] is not None else "未知"
        print(
            f"  [{tag}] 餐饮发票 {idx} ({name}): 广州={gz_ok}, 餐饮={cat_ok} | "
            f"原价税合计 {raw_str} | {fee_note} | 有效金额 {eff_str}"
        )
        if inv["amount"] is not None:
            valid_invoice_amounts.append(inv["amount"])

    if orders or invoices:
        inv_total = round(sum(valid_invoice_amounts), 2) if valid_invoice_amounts else None
        ord_total = round(sum(valid_order_amounts), 2) if valid_order_amounts else None
        if inv_total is not None and ord_total is not None:
            meal_total = min(inv_total, ord_total)
            print(
                f"  [餐费核对] 有效发票总额 ¥{inv_total:.2f} vs 有效订单总额 ¥{ord_total:.2f} "
                f"-> 取最小值 ¥{meal_total:.2f}"
            )
        else:
            meal_total = inv_total if inv_total is not None else (ord_total or 0.0)
            print(f"  [餐费核对] 餐费金额 ¥{meal_total:.2f}")
        purpose = invoices[0]["purpose"] if invoices else "餐费"
        items.append((purpose, meal_total))

    # --- 2. 交通行程单 & 交通发票（晚上 23:00 后有效）---
    if traffic_trips or traffic_invoices:
        valid_trip_amounts: list[float] = []
        for idx, doc in enumerate(traffic_trips, 1):
            name = doc["path"].name
            trips = doc["trips"]
            if not trips:
                msg = f"交通行程单 {idx} ({name}): 未解析到行程明细"
                issues.append(msg)
                print(f"  [警告] {msg}")
                if doc["declared_total"] is not None:
                    valid_trip_amounts.append(doc["declared_total"])
                continue

            for t_i, t in enumerate(trips, 1):
                dt = t["datetime"]
                amt = t["amount"]
                if dt is None:
                    issues.append(f"交通行程 {idx}-{t_i} ({name}): 未识别上车时间")
                    print(f"  [警告] 交通行程 {idx}-{t_i} ({name}): 金额 ¥{amt:.2f} | 未识别上车时间")
                    valid_trip_amounts.append(amt)
                    continue

                after_2300 = _is_after_traffic_cutoff(dt)
                age = now - dt
                within_two_weeks = timedelta(0) <= age <= timedelta(days=MAX_ORDER_AGE_DAYS)
                age_days = age.total_seconds() / 86400

                parts = []
                if after_2300:
                    parts.append("23:00后上车✓")
                else:
                    parts.append("早于23:00✗")
                    issues.append(f"交通行程 {idx}-{t_i} ({name}) 上车时间 {dt.strftime('%H:%M')} 早于 23:00，无效")

                if within_two_weeks:
                    parts.append(f"距今{age_days:.1f}天(<=14天)✓")
                else:
                    parts.append(f"距今{age_days:.1f}天(超14天)✗")
                    issues.append(f"交通行程 {idx}-{t_i} ({name}) 上车时间 {dt.strftime('%Y-%m-%d')} 距今超过两周({age_days:.1f}天)")

                tag = "通过" if (after_2300 and within_two_weeks) else "不合规"
                print(f"  [{tag}] 交通行程 {idx}-{t_i} ({name}): 上车时间 {dt.strftime('%Y-%m-%d %H:%M')} | 金额 ¥{amt:.2f} | {', '.join(parts)}")
                if after_2300:
                    valid_trip_amounts.append(amt)

        valid_tinv_amounts: list[float] = []
        for idx, tinv in enumerate(traffic_invoices, 1):
            name = tinv["path"].name
            amt_str = f"¥{tinv['amount']:.2f}" if tinv["amount"] is not None else "未知"
            print(f"  [通过] 交通发票 {idx} ({name}): 有效金额 {amt_str}")
            if tinv["amount"] is not None:
                valid_tinv_amounts.append(tinv["amount"])

        tinv_total = round(sum(valid_tinv_amounts), 2) if valid_tinv_amounts else None
        trip_total = round(sum(valid_trip_amounts), 2) if valid_trip_amounts else None
        if tinv_total is not None and trip_total is not None:
            traffic_total = min(tinv_total, trip_total)
            print(
                f"  [交通核对] 交通发票总额 ¥{tinv_total:.2f} vs 有效行程总额 ¥{trip_total:.2f} "
                f"-> 取最小值 ¥{traffic_total:.2f}"
            )
        else:
            traffic_total = tinv_total if tinv_total is not None else (trip_total or 0.0)
            print(f"  [交通核对] 交通费金额 ¥{traffic_total:.2f}")
        items.append(("交通费", traffic_total))

    return items, valid_dates, issues


def build_reimbursement_pdf(
    target_dir: Path,
    output_pdf: Path,
    dept: str = "Bida",
) -> None:
    if not target_dir.is_dir():
        raise FileNotFoundError(f"目录不存在: {target_dir}")

    print(f"[0/4] 正在扫描目录: {target_dir}")

    # Find signature image
    sig_path = None
    for pat in ("signatrue.*", "signature.*"):
        matches = sorted(target_dir.glob(pat))
        if matches:
            sig_path = matches[0]
            break
    print(f"  签名文件: {sig_path or '未找到'}")

    # Find order images
    order_dir = target_dir / "orderImage"
    order_images = (
        sorted(
            p
            for p in order_dir.iterdir()
            if p.is_file() and p.suffix.lower() in IMAGE_EXTS
        )
        if order_dir.is_dir()
        else []
    )

    # Find catering invoice files (PDF or images)
    invoice_dir = target_dir / "发票"
    invoice_files = (
        sorted(
            p
            for p in invoice_dir.iterdir()
            if p.is_file() and p.suffix.lower() in INVOICE_EXTS
        )
        if invoice_dir.is_dir()
        else []
    )

    # Find traffic trip PDFs & traffic invoice files (PDF or images)
    traffic_dir = target_dir / "traffic"
    traffic_trip_pdfs, traffic_invoice_files = classify_traffic_pdfs(traffic_dir)

    # Parse order info and invoice info with logs
    print(f"\n[1/4] 正在提取订单与行程信息 (餐饮订单 {len(order_images)} 张, 交通行程单 {len(traffic_trip_pdfs)} 份)...")
    orders = [parse_order_image(p) for p in order_images]

    def _order_sort_key(o: dict) -> tuple[int, datetime, str]:
        if o["datetime"] is not None:
            return (0, o["datetime"], o["path"].name)
        if o["date"]:
            with contextlib.suppress(ValueError):
                return (0, datetime.strptime(o["date"], "%Y-%m-%d"), o["path"].name)
        return (1, datetime.max, o["path"].name)

    orders.sort(key=_order_sort_key)
    order_images = [o["path"] for o in orders]

    for info in orders:
        p = info["path"]
        dt_str = info["datetime"].strftime("%Y-%m-%d %H:%M:%S") if info["datetime"] else (info["date"] or "未识别")
        amt_str = f"¥{info['amount']:.2f}" if info["amount"] is not None else "未识别"
        print(f"  - 订单 {p.name}: 下单时间={dt_str}, 订单金额={amt_str}")

    traffic_trips = []
    for p in traffic_trip_pdfs:
        info = parse_traffic_trip_pdf(p)
        trips_desc = ", ".join(
            f"{t['datetime'].strftime('%m-%d %H:%M') if t['datetime'] else t['date']}(¥{t['amount']:.2f})"
            for t in info["trips"]
        ) or f"合计¥{info['declared_total']:.2f}"
        print(f"  - 行程单 {p.name}: {trips_desc}")
        traffic_trips.append(info)

    print(f"\n[2/4] 正在提取发票信息 (餐饮发票 {len(invoice_files)} 份, 交通发票 {len(traffic_invoice_files)} 份)...")
    invoices = []
    for p in invoice_files:
        info = parse_invoice_file(p, default_purpose="餐费")
        raw_str = f"¥{info['raw_amount']:.2f}" if info["raw_amount"] is not None else "未识别"
        print(f"  - 餐饮发票 {p.name}: 原价税合计={raw_str}, 配送费=¥{info['delivery_fee']:.2f}, 开票日期={info['date'] or '未识别'}")
        invoices.append(info)

    traffic_invoices = []
    for p in traffic_invoice_files:
        info = parse_invoice_file(p, default_purpose="交通费")
        raw_str = f"¥{info['raw_amount']:.2f}" if info["raw_amount"] is not None else "未识别"
        print(f"  - 交通发票 {p.name}: 原价税合计={raw_str}, 开票日期={info['date'] or '未识别'}")
        traffic_invoices.append(info)

    items, valid_dates, issues = check_compliance(orders, invoices, traffic_trips, traffic_invoices)

    if not valid_dates:
        valid_dates = [inv["date"] for inv in invoices if inv["date"]]

    remark = format_date_range(valid_dates)
    fill_date = parse_dir_date(target_dir.name)
    font_name = register_chinese_font()
    grand_total = round(sum(amt for _, amt in items), 2)

    print(f"\n[4/4] 正在生成 PDF: {output_pdf}")
    print(f"  明细: {items} | 合计: ¥{grand_total:.2f} | 备注: {remark or '无'}")

    writer = PdfWriter()

    # 1. Page 1: 费用报销单
    page1_bytes = draw_reimbursement_page(
        font_name=font_name,
        dept=dept,
        fill_date=fill_date,
        items=items,
        remark=remark,
        sig_path=sig_path,
    )
    for page in PdfReader(io.BytesIO(page1_bytes)).pages:
        writer.add_page(page)

    # 2. 订单截图拼接
    for order_page_bytes in draw_order_pages(order_images):
        for page in PdfReader(io.BytesIO(order_page_bytes)).pages:
            writer.add_page(page)

    # 3. 交通行程单（放在订单下面）
    for trip_pdf in traffic_trip_pdfs:
        for page in PdfReader(str(trip_pdf)).pages:
            writer.add_page(page)

    # 4. 餐饮发票（支持 PDF 与图片）
    for inv_file in invoice_files:
        _append_invoice_to_writer(writer, inv_file)

    # 5. 交通发票（支持 PDF 与图片，放在发票后面）
    for tinv_file in traffic_invoice_files:
        _append_invoice_to_writer(writer, tinv_file)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    if issues:
        print(f"\n完成生成，但存在 {len(issues)} 项合规提醒:")
        for item in issues:
            print(f"  ! {item}")
    else:
        print(f"\n全部合规校验通过，已生成 PDF: {output_pdf}")


def main() -> None:
    parser = argparse.ArgumentParser(description="生成费用报销单合并 PDF")
    parser.add_argument(
        "dir",
        nargs="?",
        default="东百+刘唐胜+2026.11.20",
        help="报销材料目录路径",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="输出 PDF 路径（默认为 <目录名>_生成.pdf）",
    )
    parser.add_argument(
        "--dept",
        default="Bida",
        help="报销部门（默认 Bida）",
    )
    args = parser.parse_args()

    target_dir = Path(args.dir)
    output_pdf = Path(args.output) if args.output else Path(f"{target_dir.name}_生成.pdf")
    build_reimbursement_pdf(target_dir, output_pdf, dept=args.dept)


if __name__ == "__main__":
    main()
