# Scripts (local)

- `html_convert.py`: Convert PDF (path/URL) to HTML via Docling.
- `html_cleanup.py`: Wash HTML (remove Word/Office/CSS cruft), normalize images into `./img/`, and optionally merge split tables + promote numbered headings.
  - Keeps Word footnote anchors (e.g. `#_ftnref1` / `#_ftn1`) and consolidates footnote definitions into an ordered list at the end.

Example (single file):

```bash
python3 html_cleanup.py html/COM-third_ndc-F.html
```

Example (all HTML in a folder):

```bash
python3 html_cleanup.py html
```

Notes:
- Output images go to `html/img/` (next to the processed HTML), with names like `<html-basename>-1.png`.
- By default it writes a one-time backup `*.html.bak` next to the original file.
