"""OCR (PyMuPDF text / pytesseract) then Ollama text model → products JSON."""
from __future__ import annotations

import json
import re
from typing import Any

import httpx
import pymupdf

OLLAMA_URL = "http://127.0.0.1:11434/api/chat"
OLLAMA_MODEL = "llama3.2"
TIMEOUT = 120.0

EXTRACT_PROMPT = """Extract product rows from the catalogue text below.
Return ONLY valid JSON with this shape:
{"products":[{"name":"string","sku":"string","watt":number|null,"price":number|null,"currency":"USD"}]}
Rules:
- Use Fake Co. catalogue lines when present
- sku like FK-XXX-NNN when available
- No markdown fences, no commentary
Text:
"""


def ocr_pdf_pages(pdf_bytes: bytes) -> list[str]:
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    pages: list[str] = []
    try:
        for page in doc:
            text = (page.get_text("text") or "").strip()
            if len(text) >= 20:
                pages.append(text)
                continue
            # Scanned / empty text layer → raster + tesseract
            try:
                import pytesseract
                from PIL import Image
            except ImportError as e:
                raise RuntimeError(
                    "pytesseract/Pillow required for scanned pages"
                ) from e
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ocr_text = (pytesseract.image_to_string(img) or "").strip()
            pages.append(ocr_text or "")
    finally:
        doc.close()
    return pages


def _parse_products_json(raw: str) -> list[dict[str, Any]]:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
    data = json.loads(raw)
    products = data.get("products", data if isinstance(data, list) else [])
    out = []
    for p in products:
        if not isinstance(p, dict):
            continue
        out.append(
            {
                "name": p.get("name") or "",
                "sku": p.get("sku"),
                "watt": p.get("watt"),
                "price": p.get("price"),
                "currency": p.get("currency") or "USD",
            }
        )
    return out


def call_ollama(page_text: str) -> list[dict[str, Any]]:
    payload = {
        "model": OLLAMA_MODEL,
        "stream": False,
        "format": "json",
        "messages": [
            {
                "role": "user",
                "content": EXTRACT_PROMPT + page_text[:8000],
            }
        ],
    }
    try:
        with httpx.Client(timeout=TIMEOUT) as client:
            r = client.post(OLLAMA_URL, json=payload)
    except httpx.ConnectError as e:
        raise RuntimeError(
            "Ollama not reachable at 127.0.0.1:11434. Start Ollama and run: ollama pull llama3.2"
        ) from e
    if r.status_code >= 400:
        raise RuntimeError(f"Ollama error {r.status_code}: {r.text[:300]}")
    body = r.json()
    content = body.get("message", {}).get("content") or body.get("response") or ""
    return _parse_products_json(content)


def extract_pdf_ocr_llm(pdf_bytes: bytes) -> dict[str, Any]:
    pages = ocr_pdf_pages(pdf_bytes)
    if not any(p.strip() for p in pages):
        raise RuntimeError("OCR produced empty text for all pages")
    merged: list[dict[str, Any]] = []
    seen: set[str] = set()
    page_logs = []
    for i, text in enumerate(pages):
        if not text.strip():
            page_logs.append({"page": i + 1, "products": 0, "note": "empty ocr"})
            continue
        products = call_ollama(text)
        page_logs.append({"page": i + 1, "products": len(products)})
        for p in products:
            key = (p.get("sku") or "") + "|" + (p.get("name") or "")
            if key in seen:
                continue
            seen.add(key)
            merged.append(p)
    return {
        "products": merged,
        "mode": "ocr_llm",
        "model": OLLAMA_MODEL,
        "pages": page_logs,
    }
