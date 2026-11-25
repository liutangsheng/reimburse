from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

PAGE_W, PAGE_H = 595.2, 419.52  # A5 landscape (matches 费用报销单.docx & reference PDF)
CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
MIN_ORDER_TIME = (17, 30)    # 餐饮订单 17:30 以后下单才有效
MIN_TRAFFIC_TIME = (23, 0)   # 交通行程 23:00 (晚上11:00) 以后才有效
MAX_ORDER_AGE_DAYS = 14      # 下单/出行时间与当前时间不能超过两周
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
INVOICE_EXTS = {".pdf"} | IMAGE_EXTS
ORDER_DIR_NAMES = ("订单", "orders", "order")
INVOICE_DIR_NAMES = ("发票", "invoices", "invoice")


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


def parse_dir_date(dir_name: str) -> tuple[int, int, int]:
    """Extract YYYY.M.D date from directory name, falling back to today."""
    m = re.search(r"(20\d{2})\.(\d{1,2})\.(\d{1,2})", dir_name)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    now = datetime.now()
    return now.year, now.month, now.day


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
