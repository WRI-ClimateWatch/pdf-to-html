import sys
import numpy as np
from pathlib import Path
from pdf2image import convert_from_path
import easyocr


def ocr_pdf(pdf_path: str, output_format: str = "html", lang: list = ["en"], dpi: int = 200):
    """
    OCR a PDF with EasyOCR and save as clean HTML or Markdown.

    Args:
        pdf_path:      Path to your PDF file
        output_format: 'html' or 'md'
        lang:          ['en'] for English, ['ch_sim','en'] for Chinese+English
        dpi:           150 (fast) to 300 (quality)
    """
    pdf_path = Path(pdf_path)
    output_path = pdf_path.with_suffix("." + ("html" if output_format == "html" else "md"))

    print(f"Loading OCR reader (first run downloads models, ~200MB)...")
    reader = easyocr.Reader(lang, gpu=False)  # gpu=False is safe for Mac

    print(f"Loading PDF: {pdf_path}")
    pages = convert_from_path(pdf_path, dpi=dpi)
    total = len(pages)
    print(f"Found {total} pages")

    page_texts = []
    for i, page_img in enumerate(pages, start=1):
        print(f"  OCR page {i}/{total}...", end="\r")

        result = reader.readtext(np.array(page_img))

        lines = []
        for (bbox, text, confidence) in result:
            if confidence > 0.5:
                lines.append(text.strip())

        page_texts.append((i, lines))

    print(f"\nWriting {output_format.upper()} to {output_path}")

    if output_format == "html":
        _write_html(page_texts, output_path, pdf_path.name)
    else:
        _write_md(page_texts, output_path, pdf_path.name)

    print(f"Done! → {output_path}")


def _write_html(page_texts, output_path, title):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <style>
    body {{ font-family: Georgia, serif; max-width: 800px; margin: 40px auto; padding: 0 24px; color: #222; line-height: 1.7; }}
    h1 {{ font-size: 1.4em; color: #555; border-bottom: 2px solid #eee; padding-bottom: 8px; margin-top: 48px; }}
    p {{ margin: 0.6em 0; }}
    .page-divider {{ border: none; border-top: 1px dashed #ccc; margin: 40px 0; }}
    .toc {{ background: #f9f9f9; padding: 16px 24px; border-radius: 6px; margin-bottom: 40px; }}
    .toc a {{ display: block; color: #444; text-decoration: none; margin: 4px 0; }}
    .toc a:hover {{ text-decoration: underline; }}
  </style>
</head>
<body>
  <h1 style="font-size:1.8em; border-bottom: 3px solid #333;">{title}</h1>
  <div class="toc">
    <strong>Pages</strong><br>
    {''.join(f'<a href="#page-{i}">Page {i}</a>' for i, _ in page_texts)}
  </div>
"""
    for i, lines in page_texts:
        html += f'  <hr class="page-divider">\n  <h1 id="page-{i}">Page {i}</h1>\n'
        for line in lines:
            escaped = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            html += f"  <p>{escaped}</p>\n"

    html += "</body>\n</html>"
    output_path.write_text(html, encoding="utf-8")


def _write_md(page_texts, output_path, title):
    md = f"# {title}\n\n"
    for i, lines in page_texts:
        md += f"## Page {i}\n\n"
        md += "\n\n".join(lines)
        md += "\n\n---\n\n"
    output_path.write_text(md, encoding="utf-8")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 ocr_pdf.py your_file.pdf [html|md] [dpi]")
        print("  format: html (default) or md")
        print("  dpi:    200 (default), use 150 for speed or 300 for quality")
        sys.exit(1)

    pdf = sys.argv[1]
    fmt = sys.argv[2] if len(sys.argv) > 2 else "html"
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 200

    ocr_pdf(pdf, fmt, dpi=dpi)