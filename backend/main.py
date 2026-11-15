"""Mock AI document pipeline: extract → review → confirm → SQLite."""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from pydantic import BaseModel, Field

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample"
DB_PATH = ROOT / "data" / "demo.db"
FRONTEND = ROOT / "frontend"
TOOLS = ROOT / "tools"

app = FastAPI(title="ai-document-pipeline-demo")


def get_db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            sku TEXT,
            watt REAL,
            price REAL,
            currency TEXT,
            source TEXT
        )
        """
    )
    conn.commit()
    return conn


def load_mock_products() -> list[dict[str, Any]]:
    xlsx = SAMPLE / "fake_products.xlsx"
    if xlsx.exists():
        wb = load_workbook(xlsx, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            return []
        headers = [str(h) for h in rows[0]]
        out = []
        for row in rows[1:]:
            item = {headers[i]: row[i] for i in range(len(headers))}
            out.append(
                {
                    "name": item.get("name"),
                    "sku": item.get("sku"),
                    "watt": item.get("watt"),
                    "price": item.get("price"),
                    "currency": item.get("currency") or "USD",
                }
            )
        return out
    js = SAMPLE / "fake_products.json"
    if js.exists():
        return json.loads(js.read_text(encoding="utf-8")).get("products", [])
    return []


class ProductIn(BaseModel):
    name: str
    sku: str | None = None
    watt: float | None = None
    price: float | None = None
    currency: str = "USD"


class ConfirmBody(BaseModel):
    products: list[ProductIn] = Field(default_factory=list)
    source: str = "mock"


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


def _mock_payload() -> dict[str, Any]:
    products = load_mock_products()
    if not products:
        raise HTTPException(500, "No sample products. Run scripts/generate_sample_data.py")
    return {"products": products, "mode": "mock"}


@app.get("/api/extract/mock")
def extract_mock() -> dict[str, Any]:
    return _mock_payload()


@app.post("/api/extract")
async def extract(file: UploadFile = File(...)) -> dict[str, Any]:
    """Mock extract: accepts upload for UX but ignores bytes."""
    _ = await file.read()
    return _mock_payload()


@app.post("/api/extract/ocr-llm")
async def extract_ocr_llm(file: UploadFile = File(...)) -> dict[str, Any]:
    """OCR each page then Ollama llama3.2 text model → products JSON."""
    raw = await file.read()
    if not raw:
        raise HTTPException(400, "Empty file")
    name = (file.filename or "").lower()
    if not name.endswith(".pdf"):
        raise HTTPException(400, "OCR+Ollama path expects a PDF upload")
    try:
        from extract_ocr_llm import extract_pdf_ocr_llm

        return extract_pdf_ocr_llm(raw)
    except RuntimeError as e:
        raise HTTPException(503, str(e)) from e
    except Exception as e:
        raise HTTPException(500, f"OCR/LLM extract failed: {e}") from e


@app.get("/api/products")
def list_products() -> dict[str, Any]:
    conn = get_db()
    rows = conn.execute("SELECT * FROM products ORDER BY id").fetchall()
    conn.close()
    return {"products": [dict(r) for r in rows]}


@app.post("/api/confirm")
def confirm(body: ConfirmBody) -> dict[str, Any]:
    if not body.products:
        raise HTTPException(400, "No products to save")
    conn = get_db()
    saved = 0
    for p in body.products:
        conn.execute(
            "INSERT INTO products (name, sku, watt, price, currency, source) VALUES (?,?,?,?,?,?)",
            (p.name, p.sku, p.watt, p.price, p.currency, body.source),
        )
        saved += 1
    conn.commit()
    conn.close()
    return {"saved": saved}


@app.delete("/api/products")
def clear_products() -> dict[str, str]:
    conn = get_db()
    conn.execute("DELETE FROM products")
    conn.commit()
    conn.close()
    return {"status": "cleared"}


def classify_row(r: dict[str, Any]) -> str:
    if str(r.get("status") or "").lower() == "error":
        return "error"
    date = str(r.get("date") or "").strip()
    amount = str(r.get("total_amount") or "").strip()
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


@app.get("/api/clean-export/preview")
def clean_preview() -> dict[str, Any]:
    raw_xlsx = SAMPLE / "fake_extraction_raw.xlsx"
    records: list[dict[str, Any]] = []
    if raw_xlsx.exists():
        wb = load_workbook(raw_xlsx, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        headers = [str(h) for h in rows[0]]
        for row in rows[1:]:
            records.append({headers[i]: row[i] for i in range(len(headers))})
    else:
        js = SAMPLE / "fake_extraction_raw.json"
        records = json.loads(js.read_text(encoding="utf-8")).get("records", [])

    buckets: dict[str, list] = {
        "ok": [],
        "null_date": [],
        "null_amount": [],
        "error": [],
    }
    for r in records:
        b = classify_row(r)
        item = dict(r)
        item["bucket"] = b
        buckets.setdefault(b, []).append(item)
    return {"buckets": buckets, "total": len(records)}


@app.post("/api/clean-export/run")
def clean_run() -> dict[str, Any]:
    from openpyxl import Workbook

    preview = clean_preview()
    out_dir = ROOT / "data" / "exports"
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for name, rows in preview["buckets"].items():
        csv_path = out_dir / f"{name}.csv"
        if rows:
            headers = list(rows[0].keys())
            lines = [",".join(headers)]
            for r in rows:
                lines.append(",".join(str(r.get(h, "")).replace(",", ";") for h in headers))
            csv_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            written.append(str(csv_path.name))

    wb = Workbook()
    first = True
    for name, rows in preview["buckets"].items():
        ws = wb.active if first else wb.create_sheet(name)
        if first:
            ws.title = name
            first = False
        if rows:
            headers = list(rows[0].keys())
            ws.append(headers)
            for r in rows:
                ws.append([r.get(h) for h in headers])
    xlsx_path = out_dir / "cleaned.xlsx"
    wb.save(xlsx_path)
    written.append(xlsx_path.name)
    return {"written": written, "dir": str(out_dir)}


@app.get("/")
def index() -> FileResponse:
    return FileResponse(FRONTEND / "index.html")


app.mount("/static", StaticFiles(directory=FRONTEND), name="static")
app.mount(
    "/tools/schema-builder",
    StaticFiles(directory=TOOLS / "schema-builder", html=True),
    name="schema_builder",
)
app.mount(
    "/tools/clean-export",
    StaticFiles(directory=TOOLS / "clean-export", html=True),
    name="clean_export",
)
app.mount("/sample", StaticFiles(directory=SAMPLE), name="sample")
