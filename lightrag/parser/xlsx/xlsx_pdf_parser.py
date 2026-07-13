"""
xlsx_pdf_parser.py

Extracts ALL content from an xlsx file for LightRAG ingestion:
  1. Regular cell/sheet data (via openpyxl)
  2. Embedded PDF objects (via zipfile + olefile), including their text
     (via PyMuPDF)

Signature matches document_routes.py's existing `_extract_xlsx(file_bytes: bytes) -> str`
so this is a drop-in replacement — see bottom of file for the one-line swap.

Install:
    pip install olefile pymupdf   (openpyxl is already a LightRAG dependency)
"""

from __future__ import annotations

import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

import olefile
from pypdf import PdfReader
import openpyxl


# ---------------------------------------------------------------------------
# 1. Pull embedded objects (PDFs, docs, etc.) out of the xlsx zip container
# ---------------------------------------------------------------------------

def find_embedded_binaries(zf: zipfile.ZipFile) -> list[dict]:
    """Returns [{"bin_path": "xl/embeddings/oleObject1.bin", "sheet": ...}, ...]"""
    results = []
    names = zf.namelist()
    bin_files = [n for n in names if n.startswith("xl/embeddings/") and n.endswith(".bin")]

    drawing_to_bins: dict[str, list[str]] = {}
    for n in names:
        if n.startswith("xl/drawings/_rels/") and n.endswith(".rels"):
            drawing_name = "xl/drawings/" + os.path.basename(n)[: -len(".rels")]
            root = ET.fromstring(zf.read(n))
            targets = []
            for rel in root:
                target = rel.attrib.get("Target", "")
                if "embeddings/" in target:
                    targets.append("xl/embeddings/" + target.split("embeddings/")[-1])
            if targets:
                drawing_to_bins[drawing_name] = targets

    sheet_for_drawing: dict[str, str] = {}
    for n in names:
        if n.startswith("xl/worksheets/_rels/") and n.endswith(".rels"):
            sheet_name = "xl/worksheets/" + os.path.basename(n)[: -len(".rels")]
            root = ET.fromstring(zf.read(n))
            for rel in root:
                target = rel.attrib.get("Target", "")
                if "drawings/" in target:
                    drawing_name = "xl/drawings/" + target.split("drawings/")[-1]
                    sheet_for_drawing[drawing_name] = sheet_name

    bin_to_sheet: dict[str, str | None] = {}
    for drawing, bins in drawing_to_bins.items():
        sheet = sheet_for_drawing.get(drawing)
        for b in bins:
            bin_to_sheet[b] = sheet

    for b in bin_files:
        results.append({"bin_path": b, "sheet": bin_to_sheet.get(b, "unknown")})
    return results


def extract_pdf_bytes_from_ole(ole_bytes: bytes) -> bytes | None:
    """Pulls the original embedded file's bytes out of an OLE container, if it's a PDF."""
    try:
        ole = olefile.OleFileIO(BytesIO(ole_bytes))
    except Exception:
        return None

    candidate_streams = []
    for stream in ole.listdir():
        stream_path = "/".join(stream)
        if stream_path in ("Package",) or stream_path.endswith("Ole10Native") \
                or "CONTENTS" in stream_path.upper():
            candidate_streams.append(stream)

    for stream in candidate_streams:
        data = ole.openstream(stream).read()
        idx = data.find(b"%PDF-")
        if idx != -1:
            return data[idx:]

    for stream in ole.listdir():
        try:
            data = ole.openstream(stream).read()
        except Exception:
            continue
        idx = data.find(b"%PDF-")
        if idx != -1:
            return data[idx:]

    return None


def extract_all_embedded_pdfs(zf: zipfile.ZipFile) -> list[dict]:
    """Returns [{"sheet": ..., "text": extracted_pdf_text}, ...] for every embedded PDF found."""
    out = []
    embedded = find_embedded_binaries(zf)
    for item in embedded:
        raw = zf.read(item["bin_path"])
        pdf_bytes = extract_pdf_bytes_from_ole(raw)
        if not pdf_bytes:
            continue  # not a PDF (could be a zip, image, doc, etc.) — silently skip
        try:
            reader = PdfReader(BytesIO(pdf_bytes))
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as e:
            text = f"[Failed to parse embedded PDF: {e}]"
        out.append({"sheet": item["sheet"], "text": text})
    return out


# ---------------------------------------------------------------------------
# 2. Pull normal cell/sheet content
# ---------------------------------------------------------------------------

def extract_sheet_text(file_bytes: bytes) -> str:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    chunks = []
    for ws in wb.worksheets:
        rows_text = []
        for row in ws.iter_rows(values_only=True):
            cells = [str(c) for c in row if c is not None]
            if cells:
                rows_text.append(" | ".join(cells))
        if rows_text:
            chunks.append(f"=== Sheet: {ws.title} ===\n" + "\n".join(rows_text))
    return "\n\n".join(chunks)


# ---------------------------------------------------------------------------
# 3. Public entry point — drop-in replacement for document_routes._extract_xlsx
# ---------------------------------------------------------------------------

def extract_xlsx_with_embedded_pdfs(file_bytes: bytes) -> str:
    """
    Extracts sheet data AND text from any embedded PDF objects.

    Same signature as document_routes.py's `_extract_xlsx(file_bytes: bytes) -> str`,
    so it can be swapped in directly.
    """
    sheet_text = extract_sheet_text(file_bytes)

    with zipfile.ZipFile(BytesIO(file_bytes)) as zf:
        pdf_blocks = extract_all_embedded_pdfs(zf)

    parts = [sheet_text]
    for i, block in enumerate(pdf_blocks):
        parts.append(
            f"=== Embedded PDF #{i + 1} (attached near sheet: {block['sheet']}) ===\n"
            f"{block['text']}"
        )
    return "\n\n".join(parts)


# ---------------------------------------------------------------------------
# CLI test entry point (unchanged behavior from standalone testing)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python xlsx_pdf_parser.py path/to/file.xlsx")
        sys.exit(1)

    path = sys.argv[1]
    with open(path, "rb") as f:
        file_bytes = f.read()

    sheet_text = extract_sheet_text(file_bytes)
    with zipfile.ZipFile(BytesIO(file_bytes)) as zf:
        pdf_blocks = extract_all_embedded_pdfs(zf)

    print(f"Sheet text: {len(sheet_text)} characters")
    print(f"Embedded PDFs found: {len(pdf_blocks)}")
    for i, block in enumerate(pdf_blocks):
        print(f"  PDF #{i + 1} (sheet: {block['sheet']}): {len(block['text'])} characters extracted")

    combined_text = extract_xlsx_with_embedded_pdfs(file_bytes)
    print(f"\nTotal combined content: {len(combined_text)} characters")

    out_path = os.path.splitext(path)[0] + "_extracted.txt"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(combined_text)
    print(f"Full extracted text written to: {out_path}")