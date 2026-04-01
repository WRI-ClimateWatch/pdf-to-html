#!/usr/bin/env python
"""
PDF -> HTML converter (minimal).

Converts local PDF files or PDF URLs to HTML using Docling.
No LLM calls, no markdown/text outputs.
"""

import argparse
import csv
import pathlib
import re
import sys
import time
from typing import List, Dict


def _safe_stem(value: str) -> str:
    """
    Make a cross-platform-safe filename stem while preserving the user's id as much as possible.
    Keeps letters/numbers plus `._-` and turns everything else into `_`.
    """
    value = value.strip()
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("._-") or "doc"


def _read_sources(input_list: pathlib.Path) -> List[Dict[str, str]]:
    if not input_list.exists():
        raise FileNotFoundError(f"Input list not found: {input_list}")

    if input_list.suffix.lower() == ".csv":
        rows: List[Dict[str, str]] = []
        with input_list.open() as f:
            reader = csv.DictReader(f)
            for i, row in enumerate(reader, 1):
                source = (
                    row.get("source")
                    or row.get("url")
                    or row.get("path")
                    or row.get("pdf")
                )
                if not source and reader.fieldnames:
                    source = row.get(reader.fieldnames[0])
                if not source:
                    raise ValueError(f"Row {i} missing source/url/path: {row}")

                doc_id = row.get("id") or row.get("doc_id") or f"doc_{i:03d}"
                rows.append({"id": doc_id, "source": source})
        return rows

    rows = []
    for i, line in enumerate(input_list.read_text().splitlines(), 1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        rows.append({"id": f"doc_{i:03d}", "source": line})
    return rows


def _download_pdf(url: str, download_dir: pathlib.Path) -> pathlib.Path:
    import httpx

    download_dir.mkdir(parents=True, exist_ok=True)
    filename = url.split("/")[-1] or f"download_{int(time.time())}.pdf"
    if not filename.lower().endswith(".pdf"):
        filename = f"{filename}.pdf"
    tmp_pdf = download_dir / filename
    response = httpx.get(url, timeout=60, follow_redirects=True)
    response.raise_for_status()
    tmp_pdf.write_bytes(response.content)
    return tmp_pdf


def _make_converter(skip_ocr: bool):
    if not skip_ocr:
        from docling.document_converter import DocumentConverter

        return DocumentConverter()

    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions

    pdf_options = PdfPipelineOptions(do_ocr=False)
    format_options = {InputFormat.PDF: PdfFormatOption(pipeline_options=pdf_options)}
    return DocumentConverter(format_options=format_options)


def _to_local_pdf(source: str, download_dir: pathlib.Path) -> pathlib.Path:
    if source.startswith(("http://", "https://")):
        return _download_pdf(source, download_dir)

    local_path = pathlib.Path(source).expanduser().resolve()
    if not local_path.exists():
        raise FileNotFoundError(f"Local PDF not found: {local_path}")
    return local_path


def _output_name(doc_id: str, source: str) -> str:
    if doc_id:
        return f"{_safe_stem(doc_id)}.html"
    stem = pathlib.Path(source).stem
    return f"{_safe_stem(stem)}.html"


def cli() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="PDF -> HTML converter")
    p.add_argument(
        "--pdf",
        default="",
        help="Single PDF source (local path or URL)",
    )
    p.add_argument(
        "--input-list",
        default="",
        help="CSV or TXT file with PDF sources (use for batch)",
    )
    p.add_argument(
        "--out-dir",
        default="html",
        help="Directory for converted HTML files",
    )
    p.add_argument(
        "--download-dir",
        default="downloads",
        help="Directory for temporary downloaded PDFs",
    )
    p.add_argument(
        "--id",
        default="",
        help="Optional id for --pdf mode (used in output filename)",
    )
    p.add_argument(
        "--skip-ocr",
        action="store_true",
        help="Disable OCR for faster/lower-memory conversion on text-based PDFs",
    )
    return p.parse_args()


def main() -> None:
    args = cli()

    if not args.pdf and not args.input_list:
        raise ValueError("Use either --pdf or --input-list.")
    if args.pdf and args.input_list:
        raise ValueError("Use only one of --pdf or --input-list.")

    out_dir = pathlib.Path(args.out_dir)
    download_dir = pathlib.Path(args.download_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.pdf:
        doc_id = args.id or pathlib.Path(args.pdf).stem
        rows = [{"id": doc_id, "source": args.pdf}]
    else:
        rows = _read_sources(pathlib.Path(args.input_list))

    if not rows:
        print("No PDF sources found.", file=sys.stderr)
        sys.exit(1)

    converter = _make_converter(args.skip_ocr)

    for i, row in enumerate(rows, 1):
        doc_id = row["id"]
        source = row["source"]
        try:
            local_pdf = _to_local_pdf(source, download_dir)
            result = converter.convert(local_pdf)
            html_name = _output_name(doc_id, source)
            html_path = out_dir / html_name
            result.document.save_as_html(html_path)
            print(f"[{i}/{len(rows)}] OK: {doc_id} -> {html_path}")
        except Exception as e:
            print(f"[{i}/{len(rows)}] FAIL: {doc_id} - {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
