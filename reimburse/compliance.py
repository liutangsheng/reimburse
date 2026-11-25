from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

from reimburse.config import (
    IMAGE_EXTS,
    INVOICE_DIR_NAMES,
    INVOICE_EXTS,
    MAX_ORDER_AGE_DAYS,
    MIN_ORDER_TIME,
    MIN_TRAFFIC_TIME,
    ORDER_DIR_NAMES,
)


def _is_after_traffic_cutoff(dt: datetime) -> bool:
    """Valid if >= 23:00 (or early morning < 05:00)."""
    return (dt.hour, dt.minute) >= MIN_TRAFFIC_TIME or dt.hour < 5


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

    Each subdirectory containing 订单/ or 发票/ is treated as a reimbursement category.
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
    return categories


def _collect_category_files(cat: dict) -> tuple[list[Path], list[Path], list[Path]]:
    """Collect (order_images, order_pdfs, invoice_files) for a category."""
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
