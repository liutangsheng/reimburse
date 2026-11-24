from __future__ import annotations

import argparse
import contextlib
import io
import os
import re
import shutil
import subprocess
import sys
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


def _subprocess_kwargs() -> dict:
    """Suppress extra console windows when invoking subprocesses on Windows."""
    if sys.platform == "win32":
        return {"creationflags": subprocess.CREATE_NO_WINDOW}
    return {}


def _find_tesseract() -> str:
    """Locate tesseract binary from bundled package, local vendor dir, PATH, or system install."""
    exe_name = "tesseract.exe" if sys.platform == "win32" else "tesseract"
    bundled_roots: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundled_roots.append(Path(meipass))
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        bundled_roots.extend([exe_dir, exe_dir / "_internal"])
    bundled_roots.append(Path(__file__).resolve().parent / "vendor")

    for root in bundled_roots:
        candidate = root / "tesseract" / exe_name
        if candidate.exists():
            return str(candidate)

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


def _find_tessdata_dir(tess_exe: str) -> str | None:
    """Locate bundled tessdata directory next to tesseract binary if present."""
    tess_path = Path(tess_exe)
    if tess_path.is_file():
        sibling = tess_path.parent / "tessdata"
        if sibling.is_dir():
            return str(sibling)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundled = Path(meipass) / "tesseract" / "tessdata"
        if bundled.is_dir():
            return str(bundled)
    return None


def _has_tesseract_chi_sim() -> bool:
    """Check whether tesseract has the chi_sim language pack installed."""
    global _TESS_HAS_CHI_SIM
    if _TESS_HAS_CHI_SIM is None:
        try:
            tess_exe = _find_tesseract()
            cmd = [tess_exe, "--list-langs"]
            tessdata_dir = _find_tessdata_dir(tess_exe)
            if tessdata_dir:
                cmd.extend(["--tessdata-dir", tessdata_dir])
            res = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
                **_subprocess_kwargs(),
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
    tess_exe = _find_tesseract()
    cmd = [tess_exe, str(tmp_path), "stdout", "-l", lang]
    tessdata_dir = _find_tessdata_dir(tess_exe)
    if tessdata_dir:
        cmd.extend(["--tessdata-dir", tessdata_dir])
    if psm is not None:
        cmd.extend(["--psm", str(psm)])
    try:
        img.save(tmp_path)
        res = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            **_subprocess_kwargs(),
        )
        return res.stdout or ""
    finally:
        with contextlib.suppress(OSError):
            tmp_path.unlink(missing_ok=True)


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

        # Upscale bottom 28% (Meituan 合计 ¥xx.x) and middle 35% (Ele.me 实付¥xx.x) for accurate decimal OCR
        bottom = gray.crop((0, int(gray.height * 0.72), gray.width, gray.height))
        bottom_3x = bottom.resize((bottom.width * 3, bottom.height * 3), Image.Resampling.LANCZOS)
        bottom_text = _ocr_image(bottom_3x, psm=6, work_dir=work_dir)

        mid = gray.crop((0, int(gray.height * 0.35), gray.width, int(gray.height * 0.72)))
        mid_2x = mid.resize((mid.width * 2, mid.height * 2), Image.Resampling.LANCZOS)
        mid_text = _ocr_image(mid_2x, psm=6, work_dir=work_dir)

        order_dt, date_only = _extract_order_datetime([full_text, mid_text, bottom_text])
        amount = _extract_order_amount([bottom_text, mid_text, full_text])

        # Fallback with chi_sim+eng if full datetime or amount was missed
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
        # Fallback for screen-photo / uneven-lighting invoices if key fields were missed
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
            # If multiple invoices appear on the same PDF page, split by 发票号码
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
            # Fallback for scanned image-only PDF page
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
    """Format sorted unique YYYY-MM-DD dates into 'MM.DD 号加班' or 'MM.DD/MM.DD 号加班'."""
    if not dates:
        return ""
    sorted_dates = sorted(set(dates))
    fmt = lambda d: f"{int(d[5:7])}.{int(d[8:10])}"
    return f"{'/'.join(fmt(d) for d in sorted_dates)} 号加班"


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
    attachment_pages: int = 0,
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
    if attachment_pages > 0:
        c.drawCentredString(right_x + 77, sub_y, str(attachment_pages))
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

    # 备注 (vertical across rows 1..2: y from table_top - 1*row_h to table_top - 3*row_h)
    bx = (col_x[2] + col_x[3]) / 2
    draw_centered_text(c, bx, table_top - 1.6 * row_h, "备", font_name, 10)
    draw_centered_text(c, bx, table_top - 2.4 * row_h, "注", font_name, 10)

    # 领导审核 (vertical across rows 3..5)
    lx = (col_x[2] + col_x[3]) / 2
    for idx, ch in enumerate("领导审核"):
        draw_centered_text(c, lx, table_top - (3.4 + idx * 0.72) * row_h, ch, font_name, 9.5)

    # Rows 1..5: fill items (up to 5 rows)
    for row_idx, (item_name, item_amt) in enumerate(items[:5], start=1):
        draw_centered_text(c, (col_x[0] + col_x[1]) / 2, mid_y(row_idx), item_name, font_name, 10)
        draw_centered_text(c, (col_x[1] + col_x[2]) / 2, mid_y(row_idx), f"{item_amt:.2f}", font_name, 10)

    # Remark centered vertically & horizontally in rows 1..2 col 3
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
            # Split into two balanced lines if still wider than cell
            mid_idx = len(remark) // 2
            line1, line2 = remark[:mid_idx], remark[mid_idx:]
            draw_centered_text(c, rx, ry + font_sz * 0.65, line1, font_name, font_sz)
            draw_centered_text(c, rx, ry - font_sz * 0.65, line2, font_name, font_sz)

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


def _set_cell_font(run, size_pt: float = 10.0, bold: bool = False, name: str = "宋体") -> None:
    """Apply consistent Chinese font styling to a python-docx Run."""
    from docx.oxml.ns import qn
    from docx.shared import Pt

    run.font.name = name
    run.font.size = Pt(size_pt)
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        from docx.oxml import OxmlElement

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


def _render_pdf_pages_to_pngs(pdf_path: Path, scale: float = 2.5) -> list[io.BytesIO]:
    """Render each page of a PDF file into an in-memory PNG buffer via pypdfium2."""
    import pypdfium2 as pdfium

    buffers: list[io.BytesIO] = []
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for page in pdf:
            bitmap = page.render(scale=scale)
            pil_img = bitmap.to_pil().convert("RGB")
            buf = io.BytesIO()
            pil_img.save(buf, format="PNG")
            buf.seek(0)
            buffers.append(buf)
    finally:
        pdf.close()
    return buffers


def _image_to_rgb_png(img_path: Path) -> tuple[io.BytesIO, int, int]:
    """Load an image file (compositing alpha over white) into an RGB PNG buffer and return (buf, w, h)."""
    with Image.open(img_path) as raw:
        if raw.mode in ("RGBA", "LA") or (raw.mode == "P" and "transparency" in raw.info):
            rgba = raw.convert("RGBA")
            bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
            im = Image.alpha_composite(bg, rgba).convert("RGB")
        else:
            im = raw.convert("RGB")
        iw, ih = im.size
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        buf.seek(0)
        return buf, iw, ih


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

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("费 用 报 销 单")
    r_title.underline = WD_UNDERLINE.DOUBLE
    _set_cell_font(r_title, size_pt=20, bold=False)

    # Sub-header
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

    # Table (8 rows x 4 cols)
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

    # Header
    _format_docx_cell(table.cell(0, 0), "用      途", 10)
    _format_docx_cell(table.cell(0, 1), "金额（元）", 10)

    # Merge 备注 (rows 1..2) and 领导审核 (rows 3..5)
    c_bz_lbl = table.cell(1, 2).merge(table.cell(2, 2))
    _format_docx_cell(c_bz_lbl, "备\n注", 10)
    c_bz_val = table.cell(1, 3).merge(table.cell(2, 3))
    _format_docx_cell(c_bz_val, remark, 10)

    c_ld_lbl = table.cell(3, 2).merge(table.cell(5, 2))
    _format_docx_cell(c_ld_lbl, "领\n导\n审\n核", 9.5)
    table.cell(3, 3).merge(table.cell(5, 3))

    # Items (rows 1..5)
    for row_idx, (item_name, item_amt) in enumerate(items[:5], start=1):
        _format_docx_cell(table.cell(row_idx, 0), item_name, 10)
        _format_docx_cell(table.cell(row_idx, 1), f"{item_amt:.2f}", 10)

    # Row 6: 合计
    _format_docx_cell(table.cell(6, 0), "合      计", 10)
    _format_docx_cell(table.cell(6, 1), f"{total_amount:.2f}", 10)
    table.cell(6, 2).merge(table.cell(6, 3))

    # Row 7: 金额大写 + 原借款 + 应退余款
    c_r7 = table.cell(7, 0).merge(table.cell(7, 3))
    units = amount_to_chinese_units(total_amount)
    upper_parts = []
    for val, unit in units:
        mark = "⊗" if val == "ⓧ" else (val or "  ")
        upper_parts.append(f" {mark} {unit}")
    r7_text = f"金额大写:{''.join(upper_parts)}   |   原借款: ____ 元   |   应退余款: ____ 元"
    _format_docx_cell(c_r7, r7_text, 9.0, align="left")

    # Footer with signature
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

    # 2. 订单截图（每页最多 3 张横向排列）
    avail_w = PAGE_W - 104.0  # matching 52pt left/right margin
    avail_h = PAGE_H - 72.0   # matching 36pt top/30pt bottom margin
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

    # 3. 订单/行程单 PDF（每页渲染为高清图片插入）
    for pdf_path in order_pdfs or []:
        for png_buf in _render_pdf_pages_to_pngs(pdf_path):
            _add_docx_fullpage_image(doc, png_buf, avail_w, avail_h)

    # 4. 发票（图片直接插入，PDF 每页渲染为高清图片插入）
    for inv_path in invoice_files or []:
        if inv_path.suffix.lower() == ".pdf":
            for png_buf in _render_pdf_pages_to_pngs(inv_path):
                _add_docx_fullpage_image(doc, png_buf, avail_w, avail_h)
        else:
            buf, _, _ = _image_to_rgb_png(inv_path)
            _add_docx_fullpage_image(doc, buf, avail_w, avail_h)

    doc.save(str(output_docx))


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


ORDER_DIR_NAMES = ("订单", "orderimage", "orders", "order")
INVOICE_DIR_NAMES = ("发票", "invoices", "invoice")


def _find_named_subdir(parent: Path, candidates: tuple[str, ...]) -> Path | None:
    """Find a direct subdirectory matching any candidate name (case-insensitive)."""
    if not parent.is_dir():
        return None
    by_lower = {
        d.name.lower(): d
        for d in parent.iterdir()
        if d.is_dir() and not d.name.startswith(".")
    }
    for name in candidates:
        if name.lower() in by_lower:
            return by_lower[name.lower()]
    return None


def _classify_category_kind(cat_name: str) -> str:
    """Classify category into 'catering', 'traffic', or 'general' based on folder name."""
    low = cat_name.lower()
    if any(k in low for k in ("餐", "饭", "meal", "catering", "food")):
        return "catering"
    if any(k in low for k in ("交通", "打车", "出行", "车费", "traffic", "taxi")):
        return "traffic"
    return "general"


def discover_categories(target_dir: Path) -> list[dict]:
    """Discover reimbursement categories under target_dir.

    Preferred structure: each subdirectory is a category containing fixed 订单/ and 发票/ subdirs.
    Also falls back to legacy root-level orderImage/, 发票/, traffic/ layout if present.
    """
    subdirs = sorted(
        (d for d in target_dir.iterdir() if d.is_dir() and not d.name.startswith(".")),
        key=lambda d: d.name,
    )
    categories: list[dict] = []
    for d in subdirs:
        ord_dir = _find_named_subdir(d, ORDER_DIR_NAMES)
        inv_dir = _find_named_subdir(d, INVOICE_DIR_NAMES)
        if ord_dir is not None or inv_dir is not None:
            categories.append({
                "name": d.name,
                "kind": _classify_category_kind(d.name),
                "order_dir": ord_dir,
                "invoice_dir": inv_dir,
            })

    if categories:
        return categories

    # Legacy fallback: root-level orderImage/, 发票/, traffic/
    legacy_ord = _find_named_subdir(target_dir, ("orderimage", "订单"))
    legacy_inv = _find_named_subdir(target_dir, ("发票",))
    if legacy_ord is not None or legacy_inv is not None:
        categories.append({
            "name": "餐费",
            "kind": "catering",
            "order_dir": legacy_ord,
            "invoice_dir": legacy_inv,
        })
    legacy_traffic = _find_named_subdir(target_dir, ("traffic",))
    if legacy_traffic is not None:
        categories.append({
            "name": "交通费",
            "kind": "traffic",
            "order_dir": legacy_traffic,
            "invoice_dir": legacy_traffic,
            "is_legacy_traffic": True,
        })
    return categories


def _collect_category_files(cat: dict) -> tuple[list[Path], list[Path], list[Path]]:
    """Collect (order_images, order_pdfs, invoice_files) for a category."""
    if cat.get("is_legacy_traffic"):
        trip_pdfs, inv_files = classify_traffic_pdfs(cat["order_dir"])
        return [], trip_pdfs, inv_files

    order_images: list[Path] = []
    order_pdfs: list[Path] = []
    ord_dir: Path | None = cat.get("order_dir")
    if ord_dir and ord_dir.is_dir():
        for p in sorted(ord_dir.iterdir()):
            if not p.is_file():
                continue
            ext = p.suffix.lower()
            if ext in IMAGE_EXTS:
                order_images.append(p)
            elif ext == ".pdf":
                order_pdfs.append(p)

    invoice_files: list[Path] = []
    inv_dir: Path | None = cat.get("invoice_dir")
    if inv_dir and inv_dir.is_dir():
        for p in sorted(inv_dir.iterdir()):
            if p.is_file() and p.suffix.lower() in INVOICE_EXTS:
                invoice_files.append(p)

    return order_images, order_pdfs, invoice_files


def _check_category_compliance(
    cat_result: dict,
    now: datetime,
    valid_dates: list[str],
    issues: list[str],
) -> tuple[str, float]:
    """Validate one category's orders and invoices, returning (category_name, reimbursed_amount)."""
    cat_name = cat_result["name"]
    kind = cat_result["kind"]
    orders = cat_result["orders"]
    order_pdfs = cat_result["order_pdfs"]
    invoices = cat_result["invoices"]

    valid_order_amounts: list[float] = []

    # 1. Order images
    for idx, o in enumerate(orders, 1):
        name = o["path"].name
        dt = o["datetime"]
        amt = o["amount"]
        amt_str = f"¥{amt:.2f}" if amt is not None else "未识别"

        if kind == "general":
            dt_str = dt.strftime("%Y-%m-%d %H:%M:%S") if dt else (o["date"] or "未识别")
            print(f"  [通过] {cat_name}订单 {idx} ({name}): 时间 {dt_str} | 金额 {amt_str}")
            if o["date"]:
                valid_dates.append(o["date"])
            if amt is not None:
                valid_order_amounts.append(amt)
            continue

        if dt is None:
            cutoff_label = "23:00" if kind == "traffic" else "17:30"
            msg = f"{cat_name}订单 {idx} ({name}): 未能识别完整下单时间，无法校验 {cutoff_label} 与两周限制"
            issues.append(msg)
            print(f"  [警告] {msg}")
            if o["date"]:
                valid_dates.append(o["date"])
            if amt is not None:
                valid_order_amounts.append(amt)
            continue

        if kind == "traffic":
            after_cutoff = _is_after_traffic_cutoff(dt)
            cutoff_ok_txt, cutoff_bad_txt, cutoff_err = "23:00后✓", "早于23:00✗", "早于 23:00"
        else:
            after_cutoff = (dt.hour, dt.minute) >= MIN_ORDER_TIME
            cutoff_ok_txt, cutoff_bad_txt, cutoff_err = "17:30后下单✓", "早于17:30✗", "早于 17:30"

        age = now - dt
        within_two_weeks = timedelta(0) <= age <= timedelta(days=MAX_ORDER_AGE_DAYS)
        age_days = age.total_seconds() / 86400

        status_parts = []
        if after_cutoff:
            status_parts.append(cutoff_ok_txt)
        else:
            status_parts.append(cutoff_bad_txt)
            issues.append(f"{cat_name}订单 {idx} ({name}) 时间 {dt.strftime('%H:%M')} {cutoff_err}，无效")

        if within_two_weeks:
            status_parts.append(f"距今{age_days:.1f}天(<=14天)✓")
        else:
            status_parts.append(f"距今{age_days:.1f}天(超14天)✗")
            issues.append(f"{cat_name}订单 {idx} ({name}) 时间 {dt.strftime('%Y-%m-%d')} 距今超过两周({age_days:.1f}天)")

        is_valid = after_cutoff and within_two_weeks
        tag = "通过" if is_valid else "不合规"
        print(f"  [{tag}] {cat_name}订单 {idx} ({name}): 时间 {dt.strftime('%Y-%m-%d %H:%M:%S')} | 金额 {amt_str} | {', '.join(status_parts)}")
        if is_valid:
            valid_dates.append(o["date"])
            if amt is not None:
                valid_order_amounts.append(amt)

    # 2. Order PDFs (e.g. ride-hailing trip sheets or general/scanned order PDFs)
    for idx, doc in enumerate(order_pdfs, 1):
        name = doc["path"].name
        trips = doc["trips"]
        if not trips:
            dt = doc.get("datetime")
            amt = doc.get("declared_total")
            amt_str = f"¥{amt:.2f}" if amt is not None else "未识别"
            if kind == "traffic" and dt is None and amt is None:
                msg = f"{cat_name}行程单 {idx} ({name}): 未解析到行程明细"
                issues.append(msg)
                print(f"  [警告] {msg}")
            elif kind == "general" or dt is None:
                dt_str = dt.strftime("%Y-%m-%d %H:%M:%S") if dt else (doc.get("date") or "未识别")
                print(f"  [通过] {cat_name}订单PDF {idx} ({name}): 时间 {dt_str} | 金额 {amt_str}")
                if doc.get("date"):
                    valid_dates.append(doc["date"])
                if amt is not None:
                    valid_order_amounts.append(amt)
            else:
                after_cutoff = (
                    _is_after_traffic_cutoff(dt)
                    if kind == "traffic"
                    else (dt.hour, dt.minute) >= MIN_ORDER_TIME
                )
                age = now - dt
                within_two_weeks = timedelta(0) <= age <= timedelta(days=MAX_ORDER_AGE_DAYS)
                tag = "通过" if (after_cutoff and within_two_weeks) else "不合规"
                print(f"  [{tag}] {cat_name}订单PDF {idx} ({name}): 时间 {dt.strftime('%Y-%m-%d %H:%M:%S')} | 金额 {amt_str}")
                if not after_cutoff:
                    issues.append(f"{cat_name}订单PDF {idx} ({name}) 时间 {dt.strftime('%H:%M')} 不符合时间要求")
                if not within_two_weeks:
                    issues.append(f"{cat_name}订单PDF {idx} ({name}) 时间 {dt.strftime('%Y-%m-%d')} 距今超过两周")
                if after_cutoff and within_two_weeks:
                    if doc.get("date"):
                        valid_dates.append(doc["date"])
                    if amt is not None:
                        valid_order_amounts.append(amt)
            continue

        for t_i, t in enumerate(trips, 1):
            dt = t["datetime"]
            amt = t["amount"]
            if dt is None:
                issues.append(f"{cat_name}行程 {idx}-{t_i} ({name}): 未识别上车时间")
                print(f"  [警告] {cat_name}行程 {idx}-{t_i} ({name}): 金额 ¥{amt:.2f} | 未识别上车时间")
                if t.get("date"):
                    valid_dates.append(t["date"])
                valid_order_amounts.append(amt)
                continue

            if kind == "general":
                print(f"  [通过] {cat_name}行程 {idx}-{t_i} ({name}): 时间 {dt.strftime('%Y-%m-%d %H:%M')} | 金额 ¥{amt:.2f}")
                if t.get("date"):
                    valid_dates.append(t["date"])
                valid_order_amounts.append(amt)
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
                issues.append(f"{cat_name}行程 {idx}-{t_i} ({name}) 上车时间 {dt.strftime('%H:%M')} 早于 23:00，无效")

            if within_two_weeks:
                parts.append(f"距今{age_days:.1f}天(<=14天)✓")
            else:
                parts.append(f"距今{age_days:.1f}天(超14天)✗")
                issues.append(f"{cat_name}行程 {idx}-{t_i} ({name}) 上车时间 {dt.strftime('%Y-%m-%d')} 距今超过两周({age_days:.1f}天)")

            tag = "通过" if (after_2300 and within_two_weeks) else "不合规"
            print(f"  [{tag}] {cat_name}行程 {idx}-{t_i} ({name}): 上车时间 {dt.strftime('%Y-%m-%d %H:%M')} | 金额 ¥{amt:.2f} | {', '.join(parts)}")
            if after_2300:
                valid_order_amounts.append(amt)
            if after_2300 and within_two_weeks and t.get("date"):
                valid_dates.append(t["date"])

    # 3. Invoices
    valid_invoice_amounts: list[float] = []
    for idx, inv in enumerate(invoices, 1):
        name = inv["path"].name
        raw_str = f"¥{inv['raw_amount']:.2f}" if inv["raw_amount"] is not None else "未知"
        eff_str = f"¥{inv['amount']:.2f}" if inv["amount"] is not None else "未知"
        fee_note = f"扣除配送费 ¥{inv['delivery_fee']:.2f}" if inv["delivery_fee"] > 0 else "无配送费"

        if kind == "catering":
            gz_ok = inv["is_guangzhou"]
            cat_ok = inv["is_catering"]
            if not (gz_ok and cat_ok):
                issues.append(f"{cat_name}发票 {idx} ({name}) 不是广州餐饮票 (广州:{gz_ok}, 餐饮:{cat_ok})")
            tag = "通过" if (gz_ok and cat_ok) else "不合规"
            print(
                f"  [{tag}] {cat_name}发票 {idx} ({name}): 广州={gz_ok}, 餐饮={cat_ok} | "
                f"原价税合计 {raw_str} | {fee_note} | 有效金额 {eff_str}"
            )
        else:
            print(
                f"  [通过] {cat_name}发票 {idx} ({name}): "
                f"原价税合计 {raw_str} | {fee_note} | 有效金额 {eff_str}"
            )
        if inv["amount"] is not None:
            valid_invoice_amounts.append(inv["amount"])

    inv_total = round(sum(valid_invoice_amounts), 2) if valid_invoice_amounts else None
    ord_total = round(sum(valid_order_amounts), 2) if valid_order_amounts else None
    if inv_total is not None and ord_total is not None:
        cat_total = min(inv_total, ord_total)
        print(
            f"  [{cat_name}核对] 有效发票总额 ¥{inv_total:.2f} vs 有效订单总额 ¥{ord_total:.2f} "
            f"-> 取最小值 ¥{cat_total:.2f}"
        )
    else:
        cat_total = inv_total if inv_total is not None else (ord_total or 0.0)
        print(f"  [{cat_name}核对] {cat_name}金额 ¥{cat_total:.2f}")

    return cat_name, cat_total


def _print_reimbursement_summary(
    cat_results: list[dict],
    items: list[tuple[str, float]],
    grand_total: float,
) -> None:
    """Print a clear final summary of order, invoice, and actual reimbursed amounts per category."""
    reimb_map = {name: amt for name, amt in items}
    print("\n================ 报销金额汇总 ================")
    for idx, cat in enumerate(cat_results):
        cat_name = cat["name"]
        orders = cat["orders"]
        order_pdfs = cat["order_pdfs"]
        invoices = cat["invoices"]

        ord_known = [o["amount"] for o in orders if o["amount"] is not None]
        ord_unknown = len(orders) - len(ord_known)
        pdf_sum = sum(
            sum(t["amount"] for t in doc["trips"])
            if doc["trips"]
            else (doc["declared_total"] or 0.0)
            for doc in order_pdfs
        )
        ord_sum = round(sum(ord_known) + pdf_sum, 2)
        ord_count = len(orders) + len(order_pdfs)
        ord_note = f"，未识别 {ord_unknown} 张" if ord_unknown else ""

        inv_raw = round(sum(i["raw_amount"] for i in invoices if i["raw_amount"] is not None), 2)
        inv_fee = round(sum(i["delivery_fee"] for i in invoices), 2)
        inv_eff = round(sum(i["amount"] for i in invoices if i["amount"] is not None), 2)
        fee_desc = f"，扣除配送费 ¥{inv_fee:.2f}" if inv_fee > 0 else ""

        if idx > 0:
            print("  --------------------------------------------")
        print(f"  【{cat_name}】")
        print(f"    订单金额: ¥{ord_sum:.2f}（共 {ord_count} 份{ord_note}）")
        print(f"    发票金额: ¥{inv_eff:.2f}（原价税合计 ¥{inv_raw:.2f}{fee_desc}，共 {len(invoices)} 份）")
        print(f"    实际报销: ¥{reimb_map.get(cat_name, 0.0):.2f}")

    print("  --------------------------------------------")
    print(f"  实际报销总金额: ¥{grand_total:.2f}")
    print("==============================================")


def build_reimbursement_pdf(
    target_dir: Path,
    output_pdf: Path,
    dept: str = "Bida",
    output_docx: Path | None = None,
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

    categories = discover_categories(target_dir)
    if not categories:
        raise ValueError(
            f"未在 {target_dir} 下找到有效的报销类目子目录（每个子目录下需包含「订单」或「发票」文件夹）"
        )
    print(f"  识别到 {len(categories)} 个报销类目: {', '.join(c['name'] for c in categories)}")

    cat_results: list[dict] = []
    all_order_images: list[Path] = []
    all_order_pdfs: list[Path] = []
    all_invoice_files: list[Path] = []

    print("\n[1/4] 正在提取各类目订单信息...")
    for cat in categories:
        cat_name = cat["name"]
        kind = cat["kind"]
        order_images, order_pdfs, invoice_files = _collect_category_files(cat)
        print(f"  > 类目【{cat_name}】: 订单图片 {len(order_images)} 张, 订单PDF {len(order_pdfs)} 份, 发票 {len(invoice_files)} 份")

        orders = [parse_order_image(p) for p in order_images]
        orders.sort(
            key=lambda o: (
                0 if (o["datetime"] or o["date"]) else 1,
                o["datetime"]
                or (datetime.strptime(o["date"], "%Y-%m-%d") if o["date"] else datetime.max),
                o["path"].name,
            )
        )
        for info in orders:
            p = info["path"]
            dt_str = info["datetime"].strftime("%Y-%m-%d %H:%M:%S") if info["datetime"] else (info["date"] or "未识别")
            amt_str = f"¥{info['amount']:.2f}" if info["amount"] is not None else "未识别"
            print(f"    - 订单 {p.name}: 时间={dt_str}, 金额={amt_str}")

        parsed_order_pdfs = []
        for p in order_pdfs:
            info = parse_order_pdf(p)
            dt_desc = (
                info["datetime"].strftime("%Y-%m-%d %H:%M:%S")
                if info.get("datetime")
                else (info.get("date") or "未识别")
            )
            amt_desc = f"¥{info['declared_total']:.2f}" if info["declared_total"] is not None else "未识别"
            trips_desc = ", ".join(
                f"{t['datetime'].strftime('%m-%d %H:%M') if t['datetime'] else t['date']}(¥{t['amount']:.2f})"
                for t in info["trips"]
            ) or f"时间={dt_desc}, 金额={amt_desc}"
            print(f"    - 订单PDF {p.name}: {trips_desc}")
            parsed_order_pdfs.append(info)

        all_order_images.extend(o["path"] for o in orders)
        all_order_pdfs.extend(order_pdfs)
        all_invoice_files.extend(invoice_files)

        cat_results.append({
            "name": cat_name,
            "kind": kind,
            "orders": orders,
            "order_pdfs": parsed_order_pdfs,
            "invoice_files": invoice_files,
            "invoices": [],
        })

    print("\n[2/4] 正在提取各类目发票信息...")
    for cat_res in cat_results:
        cat_name = cat_res["name"]
        for p in cat_res["invoice_files"]:
            info = parse_invoice_file(p, default_purpose=cat_name)
            raw_str = f"¥{info['raw_amount']:.2f}" if info["raw_amount"] is not None else "未识别"
            fee_str = f", 配送费=¥{info['delivery_fee']:.2f}" if info["delivery_fee"] > 0 else ""
            sub_str = f", 含{info['sub_count']}张" if info.get("sub_count", 1) > 1 else ""
            print(f"  - [{cat_name}] {p.name}: 原价税合计={raw_str}{fee_str}{sub_str}, 开票日期={info['date'] or '未识别'}")
            cat_res["invoices"].append(info)

    now = datetime.now()
    issues: list[str] = []
    items: list[tuple[str, float]] = []
    valid_dates: list[str] = []
    print("\n[3/4] 正在核对合规性...")
    print(f"  当前参考时间: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    for cat_res in cat_results:
        items.append(_check_category_compliance(cat_res, now, valid_dates, issues))

    if not valid_dates:
        valid_dates = [
            inv["date"]
            for cat_res in cat_results
            for inv in cat_res["invoices"]
            if inv["date"]
        ]

    remark = format_date_range(valid_dates)
    fill_date = parse_dir_date(target_dir.name)
    font_name = register_chinese_font()
    grand_total = round(sum(amt for _, amt in items), 2)

    print(f"\n[4/4] 正在生成 PDF: {output_pdf}")
    print(f"  明细: {items} | 合计: ¥{grand_total:.2f} | 备注: {remark or '无'}")

    order_page_bytes_list = draw_order_pages(all_order_images)
    order_pdf_readers = [PdfReader(str(p)) for p in all_order_pdfs]
    invoice_pdf_pages = sum(
        len(PdfReader(str(p)).pages) if p.suffix.lower() == ".pdf" else 1
        for p in all_invoice_files
    )
    attachment_pages = (
        len(order_page_bytes_list)
        + sum(len(r.pages) for r in order_pdf_readers)
        + invoice_pdf_pages
    )

    writer = PdfWriter()

    # 1. Page 1: 费用报销单
    page1_bytes = draw_reimbursement_page(
        font_name=font_name,
        dept=dept,
        fill_date=fill_date,
        items=items,
        remark=remark,
        sig_path=sig_path,
        attachment_pages=attachment_pages,
    )
    for page in PdfReader(io.BytesIO(page1_bytes)).pages:
        writer.add_page(page)

    # 2. 所有类目订单（图片拼接 + 订单/行程单 PDF）
    for order_page_bytes in order_page_bytes_list:
        for page in PdfReader(io.BytesIO(order_page_bytes)).pages:
            writer.add_page(page)
    for reader in order_pdf_readers:
        for page in reader.pages:
            writer.add_page(page)

    # 3. 所有类目发票（支持 PDF 与图片）
    for inv_file in all_invoice_files:
        _append_invoice_to_writer(writer, inv_file)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    if output_docx is not None:
        export_reimbursement_docx(
            output_docx=output_docx,
            dept=dept,
            fill_date=fill_date,
            items=items,
            remark=remark,
            sig_path=sig_path,
            order_images=all_order_images,
            order_pdfs=all_order_pdfs,
            invoice_files=all_invoice_files,
            attachment_pages=attachment_pages,
        )
        print(f"  已同步导出 Word: {output_docx}")

    if issues:
        print(f"\n完成生成，但存在 {len(issues)} 项合规提醒:")
        for item in issues:
            print(f"  ! {item}")
    else:
        print(f"\n全部合规校验通过，已生成 PDF: {output_pdf}")

    _print_reimbursement_summary(
        cat_results=cat_results,
        items=items,
        grand_total=grand_total,
    )


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
    parser.add_argument(
        "--docx",
        action="store_true",
        help="同步输出可编辑的 Word (.docx) 报销单",
    )
    args = parser.parse_args()

    target_dir = Path(args.dir)
    output_pdf = Path(args.output) if args.output else Path(f"{target_dir.name}_生成.pdf")
    output_docx = output_pdf.with_suffix(".docx") if args.docx else None
    build_reimbursement_pdf(target_dir, output_pdf, dept=args.dept, output_docx=output_docx)


if __name__ == "__main__":
    main()
