# Scripts (local)

- `html_convert.py`: Convert PDF (path/URL) to HTML via Docling.
- `html_cleanup.py`: Wash HTML (remove Word/Office/CSS cruft), normalize images into `./img/`, and optionally merge split tables + promote numbered headings.

Example (single file):

```bash
python3 script/html_cleanup.py script/html/COM-third_ndc-F.html
```

Example (all HTML in a folder):

```bash
python3 script/html_cleanup.py script/html
```

Notes:
- Output images go to `script/html/img/` (next to the processed HTML), with names like `<html-basename>-1.png`.
- By default it writes a one-time backup `*.html.bak` next to the original file.

