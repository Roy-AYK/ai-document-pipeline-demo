"""Generate synthetic PDFs, JSON fixtures, and Excel sample files into sample/."""
from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample"
SAMPLE.mkdir(parents=True, exist_ok=True)

PRODUCTS = [
    {"name": "LED Driver 30W", "sku": "FK-LD-030", "watt": 30, "price": 12.5, "currency": "USD"},
    {"name": "LED Driver 50W", "sku": "FK-LD-050", "watt": 50, "price": 18.0, "currency": "USD"},
    {"name": "LED Driver 75W", "sku": "FK-LD-075", "watt": 75, "price": 24.0, "currency": "USD"},
    {"name": "LED Driver 100W", "sku": "FK-LD-100", "watt": 100, "price": 29.9, "currency": "USD"},
    {"name": "LED Driver 150W", "sku": "FK-LD-150", "watt": 150, "price": 38.5, "currency": "USD"},
    {"name": "Panel Light 40W", "sku": "FK-PL-040", "watt": 40, "price": 22.0, "currency": "USD"},
    {"name": "Panel Light 60W", "sku": "FK-PL-060", "watt": 60, "price": 31.5, "currency": "USD"},
    {"name": "Panel Light 72W", "sku": "FK-PL-072", "watt": 72, "price": 36.0, "currency": "USD"},
    {"name": "Flood Light 100W", "sku": "FK-FL-100", "watt": 100, "price": 34.0, "currency": "USD"},
    {"name": "Flood Light 150W", "sku": "FK-FL-150", "watt": 150, "price": 45.0, "currency": "USD"},
    {"name": "Flood Light 200W", "sku": "FK-FL-200", "watt": 200, "price": 58.0, "currency": "USD"},
    {"name": "Street Light 80W", "sku": "FK-SL-080", "watt": 80, "price": 55.0, "currency": "USD"},
    {"name": "Street Light 120W", "sku": "FK-SL-120", "watt": 120, "price": 72.0, "currency": "USD"},
    {"name": "Downlight 12W", "sku": "FK-DL-012", "watt": 12, "price": 6.8, "currency": "USD"},
    {"name": "Downlight 18W", "sku": "FK-DL-018", "watt": 18, "price": 8.2, "currency": "USD"},
    {"name": "Downlight 24W", "sku": "FK-DL-024", "watt": 24, "price": 9.9, "currency": "USD"},
    {"name": "Sensor Module A", "sku": "FK-SN-A01", "watt": 2, "price": 4.5, "currency": "USD"},
    {"name": "Sensor Module B", "sku": "FK-SN-B02", "watt": 3, "price": 5.1, "currency": "USD"},
    {"name": "Sensor Module C", "sku": "FK-SN-C03", "watt": 4, "price": 6.0, "currency": "USD"},
    {"name": "Controller Hub", "sku": "FK-CH-001", "watt": 8, "price": 39.0, "currency": "USD"},
    {"name": "Controller Hub Pro", "sku": "FK-CH-002", "watt": 12, "price": 52.0, "currency": "USD"},
    {"name": "Dimmer Switch", "sku": "FK-DM-010", "watt": 1, "price": 7.5, "currency": "USD"},
]

RAW_ROWS = [
    {"request_id": "req-001", "date": "2025-03-15", "total_amount": "1200.00", "vendor": "Fake Co. Ltd", "status": "completed"},
    {"request_id": "req-002", "date": "", "total_amount": "880.50", "vendor": "Dummy Supplies", "status": "completed"},
    {"request_id": "req-003", "date": "2025-04-01", "total_amount": "", "vendor": "Fake Co. Ltd", "status": "completed"},
    {"request_id": "req-004", "date": "not-a-date", "total_amount": "100.00", "vendor": "Acme Demo", "status": "completed"},
    {"request_id": "req-005", "date": "2025-05-20", "total_amount": "abc", "vendor": "Dummy Supplies", "status": "completed"},
    {"request_id": "req-006", "date": "2025-06-01", "total_amount": "50.00", "vendor": "Fake Co. Ltd", "status": "error"},
    {"request_id": "req-007", "date": "2099-01-01", "total_amount": "10.00", "vendor": "Future Vendor", "status": "completed"},
    {"request_id": "req-008", "date": "2025-07-12", "total_amount": "250.00", "vendor": "Fake Co. Ltd", "status": "completed"},
    {"request_id": "req-009", "date": "", "total_amount": "", "vendor": "Dummy Supplies", "status": "completed"},
    {"request_id": "req-010", "date": "2025-08-01", "total_amount": "999.99", "vendor": "Fake Co. Ltd", "status": "completed"},
    {"request_id": "req-011", "date": "2025-08-15", "total_amount": "15.00", "vendor": "Acme Demo", "status": "error"},
    {"request_id": "req-012", "date": "15/08/2025", "total_amount": "40.00", "vendor": "Fake Co. Ltd", "status": "completed"},
    {"request_id": "req-013", "date": "2025-09-01", "total_amount": "N/A", "vendor": "Dummy Supplies", "status": "completed"},
    {"request_id": "req-014", "date": "2025-09-20", "total_amount": "320.10", "vendor": "Fake Co. Ltd", "status": "completed"},
    {"request_id": "req-015", "date": "", "total_amount": "75.00", "vendor": "Acme Demo", "status": "completed"},
    {"request_id": "req-016", "date": "2025-10-05", "total_amount": "110.00", "vendor": "Fake Co. Ltd", "status": "completed"},
]


def _classify(r: dict) -> str:
    if r.get("status") == "error":
        return "error"
    date = (r.get("date") or "").strip()
    amount = (r.get("total_amount") or "").strip()
    if not date:
        return "null_date"
    if not amount:
        return "null_amount"
    if date == "not-a-date" or date.startswith("2099"):
        return "error"
    try:
        float(amount)
    except ValueError:
        return "error"
    return "ok"


def write_products_xlsx() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "products"
    ws.append(["name", "sku", "watt", "price", "currency"])
    for p in PRODUCTS:
        ws.append([p["name"], p["sku"], p["watt"], p["price"], p["currency"]])
    wb.save(SAMPLE / "fake_products.xlsx")


def write_raw_xlsx() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "raw"
    ws.append(["request_id", "date", "total_amount", "vendor", "status"])
    for r in RAW_ROWS:
        ws.append([r["request_id"], r["date"], r["total_amount"], r["vendor"], r["status"]])
    wb.save(SAMPLE / "fake_extraction_raw.xlsx")


def write_clean_xlsx() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "clean"
    ws.append(["request_id", "date", "total_amount", "vendor", "status", "bucket"])
    for r in RAW_ROWS:
        ws.append([r["request_id"], r["date"], r["total_amount"], r["vendor"], r["status"], _classify(r)])
    wb.save(SAMPLE / "fake_extraction_clean.xlsx")


def write_json() -> None:
    (SAMPLE / "fake_products.json").write_text(
        json.dumps({"products": PRODUCTS}, indent=2), encoding="utf-8"
    )
    (SAMPLE / "fake_extraction_raw.json").write_text(
        json.dumps({"records": RAW_ROWS}, indent=2), encoding="utf-8"
    )


def _write_pdf(path: Path, lines: list[str], title: str) -> None:
    text_lines = [f"BT /F1 12 Tf 50 750 Td ({title}) Tj"]
    for line in lines:
        safe = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        text_lines.append(f"0 -18 Td ({safe}) Tj")
    text_lines.append("ET")
    content = "\n".join(text_lines).encode("latin-1", errors="replace")
    stream = b"<< /Length %d >>\nstream\n%s\nendstream" % (len(content), content)
    objects = [
        b"1 0 obj<< /Type /Catalog /Pages 2 0 R >>endobj\n",
        b"2 0 obj<< /Type /Pages /Kids [3 0 R] /Count 1 >>endobj\n",
        b"3 0 obj<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>endobj\n",
        b"4 0 obj" + stream + b"endobj\n",
        b"5 0 obj<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>endobj\n",
    ]
    out = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for obj in objects:
        offsets.append(len(out))
        out.extend(obj)
    xref_pos = len(out)
    out.extend(b"xref\n0 %d\n" % (len(objects) + 1))
    out.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        out.extend(b"%010d 00000 n \n" % off)
    out.extend(
        b"trailer<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n"
        % (len(objects) + 1, xref_pos)
    )
    path.write_bytes(bytes(out))


def write_pdfs() -> None:
    lines1 = [
        f"SKU {p['sku']} {p['name']} USD {p['price']}" for p in PRODUCTS[:8]
    ]
    lines2 = [
        f"SKU {p['sku']} {p['name']} USD {p['price']}" for p in PRODUCTS[8:16]
    ]
    _write_pdf(SAMPLE / "fake_catalogue.pdf", lines1, "Fake Co. Ltd - Product Catalogue")
    _write_pdf(SAMPLE / "fake_catalogue_p2.pdf", lines2, "Fake Co. Ltd - Catalogue Page 2")


def main() -> None:
    write_products_xlsx()
    write_raw_xlsx()
    write_clean_xlsx()
    write_json()
    write_pdfs()
    print("Wrote sample files to", SAMPLE)


if __name__ == "__main__":
    main()
