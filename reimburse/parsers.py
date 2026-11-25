from __future__ import annotations

import contextlib
import io
import re
from datetime import datetime
from pathlib import Path

from PIL import Image
from pypdf import PdfReader

from reimburse.ocr import _has_tesseract_chi_sim, _ocr_image, _to_grayscale_on_white


def _extract_order_datetime(texts: list[str]) -> tuple[datetime | None, str | None]:
    """Extract order datetime and date string from OCR texts, tolerating minor OCR punctuation glitches."""
    dt_pat = re.compile(
        r"(20\d{2})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\s+(\d{1,2})\s*[:;.]\s*(\d{2})(?:\s*[:;.]\s*(\d{2}))?"
    )
    for txt in texts:
        for m in dt_pat.finditer(txt):
            y, mo, d, h, mi, s = (
                int(m.group(1)),
                int(m.group(2)),
                int(m.group(3)),
                int(m.group(4)),
                int(m.group(5)),
                int(m.group(6) or 0),
            )
            try:
                dt = datetime(y, mo, d, h, mi, s)
                return dt, dt.strftime("%Y-%m-%d")
            except ValueError:
                continue

    d_pat = re.compile(r"(20\d{2})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})")
    for txt in texts:
        for m in d_pat.finditer(txt):
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            try:
                dt_d = datetime(y, mo, d)
                return None, dt_d.strftime("%Y-%m-%d")
            except ValueError:
                continue
    return None, None


def _extract_order_amount(texts: list[str]) -> float | None:
    """Extract order total amount from OCR texts."""
    for txt in texts:
        yen_nums = re.findall(r"(?:[¥￥]|合计\s*[¥￥]?)\s*(\d{1,4}\.\d{1,2})", txt)
        if yen_nums:
            return float(yen_nums[-1])
    for txt in texts:
        decimals = re.findall(r"\b(\d{1,4}\.\d{1,2})\b", txt)
        if decimals:
            return float(decimals[-1])
    return None


def parse_order_image(img_path: Path) -> dict:
    """Extract order datetime and amount from an order screenshot via Tesseract OCR."""
    work_dir = img_path.parent
    with Image.open(img_path) as im:
        gray = _to_grayscale_on_white(im)
        full_text = _ocr_image(gray, work_dir=work_dir)

        bottom = gray.crop((0, int(gray.height * 0.72), gray.width, gray.height))
        bottom_3x = bottom.resize((bottom.width * 3, bottom.height * 3), Image.Resampling.LANCZOS)
        bottom_text = _ocr_image(bottom_3x, psm=6, work_dir=work_dir)

        mid = gray.crop((0, int(gray.height * 0.35), gray.width, int(gray.height * 0.72)))
        mid_2x = mid.resize((mid.width * 2, mid.height * 2), Image.Resampling.LANCZOS)
        mid_text = _ocr_image(mid_2x, psm=6, work_dir=work_dir)

        order_dt, date_only = _extract_order_datetime([full_text, mid_text, bottom_text])
        amount = _extract_order_amount([bottom_text, mid_text, full_text])

        if (order_dt is None or amount is None) and _has_tesseract_chi_sim():
            extra_texts: list[str] = []
            if order_dt is None:
                extra_texts.append(_ocr_image(mid_2x, psm=6, work_dir=work_dir, lang="chi_sim+eng"))
                extra_texts.append(_ocr_image(gray, work_dir=work_dir, lang="chi_sim+eng"))
                dt2, d2 = _extract_order_datetime(extra_texts)
                if dt2 is not None:
                    order_dt, date_only = dt2, d2
                elif date_only is None and d2 is not None:
                    date_only = d2
            if amount is None:
                bottom_chi = _ocr_image(bottom_3x, psm=6, work_dir=work_dir, lang="chi_sim+eng")
                amount = _extract_order_amount([bottom_chi, *extra_texts])

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
        nums = [float(x) for x in re.findall(r"(?<!\d)(\d+\.\d{1,2})(?!\d|%)", line)]
        if not nums:
            continue
        if len(nums) >= 2 and rate is not None and abs(nums[-1] - nums[-2] * rate) <= 0.05:
            fee_total += nums[-2] + nums[-1]
        else:
            fee_total += nums[-1]
    return round(fee_total, 2)


def _find_yen_amount(text: str) -> float | None:
    """Match either '¥56.20', '56.20¥', or OCR variants like '¥43, 00' / 'Y31. 56'."""
    m = re.search(
        r"[¥￥Y]\s*(-?\d+(?:\s*[.,]\s*\d{1,2})?)|(-?\d+(?:\s*[.,]\s*\d{1,2})?)\s*[¥￥]",
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
        gz_codes = [
            re.sub(r"\s+", "", m.group(0))
            for m in re.finditer(r"9[12]4401\d{2}\s*[0-9A-Z]{10}", full_text)
        ]
        seller_gz = [c for c in gz_codes if c != "914401055622643150"]
        if seller_gz or len(gz_codes) >= 2:
            is_guangzhou = True

    combined = f"{full_text}\n{inv_path.name}"
    compact = re.sub(r"\s+", "", combined)
    is_traffic = any(k in compact for k in ("交通", "客运", "出行")) or "didi" in compact.lower()
    is_hotel = "住宿" in compact
    is_catering = (
        any(k in compact for k in ("餐饮", "饮食服务", "美食店", "餐厅", "饭店", "餐馆", "外卖"))
        or ("生产生活服务" in compact and not is_traffic and not is_hotel)
    )

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
        texts.append(_ocr_image(gray, psm=6, work_dir=work_dir, lang="chi_sim+eng"))
    texts.append(_ocr_image(gray, work_dir=work_dir, lang="eng"))
    texts.append(_ocr_image(gray, psm=6, work_dir=work_dir, lang="eng"))
    return "\n".join(texts), has_chi


def _ocr_invoice_denoised(im: Image.Image, work_dir: Path | None = None) -> str:
    """Denoise and normalize uneven lighting / moire screen photos before OCR."""
    from PIL import ImageChops, ImageFilter, ImageOps

    gray = _to_grayscale_on_white(im)
    scale = 2 if gray.width < 2000 else 1
    if scale > 1:
        gray = gray.resize((gray.width * scale, gray.height * scale), Image.Resampling.LANCZOS)
    denoised = gray.filter(ImageFilter.MedianFilter(size=3))
    bg = denoised.filter(ImageFilter.GaussianBlur(radius=25))
    diff = ImageChops.subtract(bg, denoised)
    norm = ImageOps.invert(ImageOps.autocontrast(diff, cutoff=1))
    bw = norm.point(lambda x: 0 if x < 185 else 255)

    texts = []
    if _has_tesseract_chi_sim():
        texts.append(_ocr_image(bw, work_dir=work_dir, lang="chi_sim+eng"))
        texts.append(_ocr_image(bw, psm=6, work_dir=work_dir, lang="chi_sim+eng"))
    texts.append(_ocr_image(bw, work_dir=work_dir, lang="eng"))
    return "\n".join(texts)


def parse_invoice_image(img_path: Path, default_purpose: str = "餐费") -> dict:
    """Extract total amount, delivery fee, seller, date, and category from an invoice image."""
    with Image.open(img_path) as im:
        full_text, has_chi = _ocr_invoice_pil(im, work_dir=img_path.parent)
        res = _parse_invoice_text(
            full_text,
            img_path,
            pages=1,
            is_image=True,
            has_chi_sim=has_chi,
            default_purpose=default_purpose,
        )
        if res["date"] is None or res["raw_amount"] is None or not res["is_guangzhou"]:
            denoised_text = _ocr_invoice_denoised(im, work_dir=img_path.parent)
            res2 = _parse_invoice_text(
                denoised_text,
                img_path,
                pages=1,
                is_image=True,
                has_chi_sim=has_chi,
                default_purpose=default_purpose,
            )
            if res2["date"] is not None or res2["raw_amount"] is not None:
                return {
                    **res2,
                    "raw_amount": res2["raw_amount"] if res2["raw_amount"] is not None else res["raw_amount"],
                    "amount": res2["amount"] if res2["amount"] is not None else res["amount"],
                    "date": res2["date"] or res["date"],
                    "is_guangzhou": res2["is_guangzhou"] or res["is_guangzhou"],
                    "is_catering": res2["is_catering"] or res["is_catering"],
                }
    return res


def _aggregate_invoice_parts(
    parts: list[dict],
    pdf_path: Path,
    pages: int,
    default_purpose: str,
) -> dict:
    """Aggregate parsed invoice results from multiple pages/invoices in one PDF."""
    if not parts:
        return {
            "path": pdf_path,
            "raw_amount": None,
            "delivery_fee": 0.0,
            "amount": None,
            "date": None,
            "is_guangzhou": False,
            "is_catering": False,
            "purpose": default_purpose,
            "pages": pages,
            "sub_count": 0,
        }
    if len(parts) == 1:
        parts[0]["pages"] = pages
        parts[0]["sub_count"] = 1
        return parts[0]

    with_amt = [p for p in parts if p["raw_amount"] is not None]
    active = with_amt if with_amt else parts
    raw_amount = round(sum(p["raw_amount"] for p in with_amt), 2) if with_amt else None
    delivery_fee = round(sum(p["delivery_fee"] for p in active), 2)
    eff_vals = [p["amount"] for p in active if p["amount"] is not None]
    effective_amount = round(sum(eff_vals), 2) if eff_vals else None

    return {
        "path": pdf_path,
        "raw_amount": raw_amount,
        "delivery_fee": delivery_fee,
        "amount": effective_amount,
        "date": next((p["date"] for p in active if p["date"]), None),
        "is_guangzhou": all(p["is_guangzhou"] for p in active),
        "is_catering": all(p["is_catering"] for p in active),
        "purpose": active[0]["purpose"],
        "pages": pages,
        "sub_count": len(active),
    }


def parse_invoice_pdf(pdf_path: Path, default_purpose: str = "餐费") -> dict:
    """Extract and sum amounts, delivery fees, and metadata from a single- or multi-invoice PDF."""
    reader = PdfReader(str(pdf_path))
    parts: list[dict] = []

    for page_idx, page in enumerate(reader.pages):
        page_text = page.extract_text() or ""
        if page_text.strip():
            if len(re.findall(r"发票号码", page_text)) >= 2:
                chunks = re.split(r"(?=发票号码)", page_text)
                for chunk in chunks:
                    if chunk.strip():
                        parts.append(
                            _parse_invoice_text(
                                chunk,
                                pdf_path,
                                pages=1,
                                is_image=False,
                                has_chi_sim=True,
                                default_purpose=default_purpose,
                            )
                        )
            else:
                parts.append(
                    _parse_invoice_text(
                        page_text,
                        pdf_path,
                        pages=1,
                        is_image=False,
                        has_chi_sim=True,
                        default_purpose=default_purpose,
                    )
                )
        else:
            ocr_parts: list[str] = []
            has_chi = _has_tesseract_chi_sim()
            for img_file in getattr(page, "images", []):
                with Image.open(io.BytesIO(img_file.data)) as im:
                    t, has_chi = _ocr_invoice_pil(im, work_dir=pdf_path.parent)
                    ocr_parts.append(t)
            if not ocr_parts:
                with contextlib.suppress(Exception):
                    import pypdfium2 as pdfium

                    pdf_doc = pdfium.PdfDocument(str(pdf_path))
                    try:
                        pil_img = pdf_doc[page_idx].render(scale=2.5).to_pil()
                        t, has_chi = _ocr_invoice_pil(pil_img, work_dir=pdf_path.parent)
                        ocr_parts.append(t)
                    finally:
                        pdf_doc.close()
            if ocr_parts:
                parts.append(
                    _parse_invoice_text(
                        "\n".join(ocr_parts),
                        pdf_path,
                        pages=1,
                        is_image=True,
                        has_chi_sim=has_chi,
                        default_purpose=default_purpose,
                    )
                )

    return _aggregate_invoice_parts(parts, pdf_path, len(reader.pages), default_purpose)


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
        "datetime": None,
        "date": None,
        "pages": len(reader.pages),
    }


def parse_order_pdf(pdf_path: Path) -> dict:
    """Extract trip details or general order datetime/amount from an order PDF (text or scanned)."""
    info = parse_traffic_trip_pdf(pdf_path)
    if info["trips"]:
        return info

    reader = PdfReader(str(pdf_path))
    full_text = "\n".join(page.extract_text() or "" for page in reader.pages)
    if not full_text.strip():
        ocr_parts: list[str] = []
        has_chi = _has_tesseract_chi_sim()
        lang = "chi_sim+eng" if has_chi else "eng"
        for page in reader.pages:
            for img_file in getattr(page, "images", []):
                with Image.open(io.BytesIO(img_file.data)) as im:
                    gray = _to_grayscale_on_white(im)
                    ocr_parts.append(_ocr_image(gray, work_dir=pdf_path.parent, lang=lang))
        full_text = "\n".join(ocr_parts)

    dt, date_only = _extract_order_datetime([full_text])
    if date_only is None:
        cn_d = re.search(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", full_text)
        if cn_d:
            date_only = f"{cn_d.group(1)}-{int(cn_d.group(2)):02d}-{int(cn_d.group(3)):02d}"

    amount = info["declared_total"]
    if amount is None:
        amount = _extract_order_amount([full_text])

    info["datetime"] = dt
    info["date"] = date_only
    info["declared_total"] = amount
    return info
