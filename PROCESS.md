# PDF → HTML (Word Edit) → Clean

This is the safest workflow when you want to manually insert/adjust images in Word, but still end up with clean HTML and consistently named images.

## 1) Convert PDF to HTML (Docling)

From repo root:

```bash
python3 html_convert.py --pdf "/absolute/path/to/file.pdf" --id "COM-third_ndc-F" --out-dir html
```

If your PDF is a URL:

```bash
python3 html_convert.py --pdf "https://example.com/file.pdf" --id "COM-third_ndc-F" --out-dir html
```

Output:
- HTML goes to `html/<id>.html`

## 2) Edit in Word (insert images / manual cleanup)

1. Open `html/<id>.html` in Word.
2. Insert images / adjust layout.
3. Save back to **HTML**.

Word usually creates an adjacent asset folder next to the HTML (examples):
- `html/<id>.fld/`
- `html/<id>_files/`

The HTML will reference images from that folder, often with names like `image001.png`.

## 3) Run the cleanup tool (washer + image normalization)

This does:
- removes Word/Office cruft + CSS
- merges some split lists/tables
- promotes numbered headings (`1.` → `h2`, `1.2` → `h3`, `1.2.1` → `h4`)
- turns `Tableau ...` / `Figure ...` headings into plain bold text
- preserves Word-style footnote anchors (`#_ftnref1`, `#_ftn1`) and hoists footnote definitions into a single ordered list at the end
- copies/renames images into `html/img/` as PNGs and rewrites `<img src>` to `img/<id>-N.png`

Single file:

```bash
python3 html_cleanup.py html/<id>.html
```

All HTML files in `html/`:

```bash
python3 html_cleanup.py html
```

Notes:
- Creates a one-time backup next to each file: `html/<id>.html.bak` (disable with `--no-backup`).
- Output images are written to: `html/img/`

## 4) Review and final edits

Open the cleaned HTML in a browser and spot-check:
- headings (especially multi-level numbering like `8.5.2.`)
- merged tables (no duplicated header rows)
- bibliography list numbering (not restarting at 1 mid-way)
- image placement and sizing

If you need to make more edits in Word:
- edit `html/<id>.html` again,
- then re-run the cleanup script.
