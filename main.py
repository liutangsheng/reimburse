from __future__ import annotations

import argparse
import io
from datetime import datetime
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from reimburse.compliance import (
    _check_category_compliance,
    _collect_category_files,
    _print_reimbursement_summary,
    discover_categories,
)
from reimburse.config import (
    format_date_range,
    parse_dir_date,
    register_chinese_font,
)
from reimburse.docx_export import export_reimbursement_docx
from reimburse.ocr import _has_tesseract_chi_sim, _ocr_image, _to_grayscale_on_white
from reimburse.parsers import (
    _parse_invoice_text,
    parse_invoice_file,
    parse_invoice_image,
    parse_invoice_pdf,
    parse_order_image,
    parse_order_pdf,
    parse_traffic_trip_pdf,
)
from reimburse.pdf_export import (
    _append_invoice_to_writer,
    draw_invoice_image_page,
    draw_order_pages,
    draw_reimbursement_page,
)

__all__ = [
    "build_reimbursement_pdf",
    "discover_categories",
    "draw_invoice_image_page",
    "draw_order_pages",
    "draw_reimbursement_page",
    "export_reimbursement_docx",
    "format_date_range",
    "main",
    "parse_invoice_file",
    "parse_invoice_image",
    "parse_invoice_pdf",
    "parse_order_image",
    "parse_order_pdf",
    "parse_traffic_trip_pdf",
    "register_chinese_font",
]


def build_reimbursement_pdf(
    target_dir: Path,
    output_pdf: Path,
    dept: str = "Bida",
    output_docx: Path | None = None,
) -> None:
    """Scan category folders, run OCR & compliance checks, and export merged PDF (and optional Word)."""
    if not target_dir.is_dir():
        raise FileNotFoundError(f"目录不存在: {target_dir}")

    print(f"[0/4] 正在扫描目录: {target_dir}")

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
        print(
            f"  > 类目【{cat_name}】: 订单图片 {len(order_images)} 张, "
            f"订单PDF {len(order_pdfs)} 份, 发票 {len(invoice_files)} 份"
        )

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

    for order_page_bytes in order_page_bytes_list:
        for page in PdfReader(io.BytesIO(order_page_bytes)).pages:
            writer.add_page(page)
    for reader in order_pdf_readers:
        for page in reader.pages:
            writer.add_page(page)

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
        default="示例公司+张三+2026.11.20",
        help="报销材料目录路径",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="输出 PDF 路径（默认为 <目录名>.pdf）",
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
    output_pdf = Path(args.output) if args.output else Path(f"{target_dir.name}.pdf")
    output_docx = output_pdf.with_suffix(".docx") if args.docx else None
    build_reimbursement_pdf(target_dir, output_pdf, dept=args.dept, output_docx=output_docx)


if __name__ == "__main__":
    main()
