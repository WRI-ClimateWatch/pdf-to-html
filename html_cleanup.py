#!/usr/bin/env python3
"""
HTML cleanup + image normalization.

Primary goals:
- "Washer": strip Word/Office cruft and CSS, keeping only semantic structure.
- Normalize images into a shared ./img folder with stable, doc-prefixed PNG names.
- Optional structural post-processing:
  - Merge adjacent tables likely split across pages.
  - Promote numbered section titles to heading levels.
"""

from __future__ import annotations

import argparse
import html
import os
import pathlib
import re
import shutil
from dataclasses import dataclass
from html.parser import HTMLParser
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


VOID_TAGS = {
    "br",
    "hr",
    "img",
    "meta",
}


ALLOWED_TAGS = {
    "html",
    "head",
    "body",
    "title",
    "meta",
    "p",
    "br",
    "hr",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "ul",
    "ol",
    "li",
    "table",
    "thead",
    "tbody",
    "tfoot",
    "tr",
    "th",
    "td",
    "figure",
    "figcaption",
    "a",
    "img",
    "pre",
    "code",
    "blockquote",
    "strong",
    "em",
    "sup",
    "sub",
}


ALLOWED_ATTRS: Dict[str, set[str]] = {
    "meta": {"charset", "content", "name"},
    # Keep Word footnote anchors like:
    #   <a name="_ftnref1" id="_ftnref1" href="#_ftn1"><sup>1</sup></a>
    # and footnote definition targets like:
    #   <a name="_ftn1" id="_ftn1" href="#_ftnref1">[1]</a>
    "a": {"href", "id", "name"},
    "img": {"src", "alt", "width", "height"},
    "td": {"colspan", "rowspan"},
    "th": {"colspan", "rowspan"},
    "ol": {"start"},
}


SKIP_CONTENT_TAGS = {
    "style",
    "script",
}


WORD_NAMESPACE_PREFIXES = ("o:", "v:", "w:", "m:")


_RE_MULTI_WS = re.compile(r"\s+")
# Numbered section titles like:
#   1. Title
#   1.2 Title
#   1.2.1. Title
# Allow "2.Title" (no space) as well (common in some exports).
_RE_NUMBERED = re.compile(r"^\s*(\d+(?:\.\d+)*)(\.)?\s*(\S.*\S|\S)\s*$")
_RE_PAGE_NUM = re.compile(r"^\s*(\d+|[ivxlcdm]+)\s*$", re.IGNORECASE)

HEADING_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6"}

_RE_FTN_REF = re.compile(r"^_(?:ftn|edn)ref(\d+)$", re.IGNORECASE)
_RE_FTN_DEF = re.compile(r"^_(?:ftn|edn)(\d+)$", re.IGNORECASE)


def _safe_stem(value: str) -> str:
    value = value.strip()
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("._-") or "doc"


def _norm_text(value: str) -> str:
    value = html.unescape(value)
    value = value.replace("\u00A0", " ")
    value = _RE_MULTI_WS.sub(" ", value).strip()
    return value


class Washer(HTMLParser):
    """
    Streaming HTML "washer" that:
    - drops styles/scripts/comments
    - strips inline styles/classes and Word-only tags
    - normalizes b/i -> strong/em
    - emits XML-friendly HTML (self-closing void tags, balanced tag stack)
    """

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._out: List[str] = []
        self._emit_stack: List[str] = []
        self._skip_depth: int = 0
        self._skip_tag: Optional[str] = None

    def get_output(self) -> str:
        # Close any still-open tags
        while self._emit_stack:
            tag = self._emit_stack.pop()
            if tag not in VOID_TAGS:
                self._out.append(f"</{tag}>")
        return "".join(self._out)

    def handle_decl(self, decl: str) -> None:
        # Drop doctype/decl; we'll output plain HTML.
        return

    def handle_comment(self, data: str) -> None:
        # Keep the useful part of Office conditional comments when they contain <img ...>,
        # otherwise drop comments entirely.
        s = data.strip()
        low = s.lower()
        if not low.startswith("[if"):
            return
        if "<img" not in low:
            return

        # Typical shape:
        #   [if !vml]><img ...><![endif]
        # Extract the inner HTML between "]>" and "<![endif]".
        start = low.find("]>")
        if start == -1:
            return
        end = low.rfind("<![endif")
        inner = s[start + 2 : end if end != -1 else None].strip()
        if not inner:
            return
        nested = Washer()
        nested.feed(inner)
        self._out.append(nested.get_output())
        return

    def _should_skip_tag(self, tag: str) -> bool:
        if tag in SKIP_CONTENT_TAGS:
            return True
        return False

    def _map_tag(self, tag: str) -> str:
        tag = tag.lower()
        if tag == "b":
            return "strong"
        if tag == "i":
            return "em"
        return tag

    def _filter_attrs(self, tag: str, attrs: Sequence[Tuple[str, Optional[str]]]) -> List[Tuple[str, str]]:
        allowed = ALLOWED_ATTRS.get(tag, set())
        out: List[Tuple[str, str]] = []
        for k, v in attrs:
            if v is None:
                continue
            k = k.lower()
            if k in {"style", "class", "lang"}:
                continue
            if k.startswith("on"):
                continue
            if k in allowed:
                out.append((k, v))
        return out

    def _emit_start(self, tag: str, attrs: List[Tuple[str, str]]) -> None:
        if tag not in ALLOWED_TAGS:
            return
        if tag == "meta":
            name_val = next((v for k, v in attrs if k == "name"), "").lower()
            if name_val in {"progid", "generator", "originator"}:
                return
        attr_txt = ""
        if attrs:
            parts = [f'{k}="{html.escape(v, quote=True)}"' for k, v in attrs]
            attr_txt = " " + " ".join(parts)
        if tag in VOID_TAGS:
            self._out.append(f"<{tag}{attr_txt}/>")
        else:
            self._out.append(f"<{tag}{attr_txt}>")
            self._emit_stack.append(tag)

    def _emit_end(self, tag: str) -> None:
        if tag in VOID_TAGS:
            return
        if tag not in ALLOWED_TAGS:
            return
        if not self._emit_stack:
            return
        # HTML in the wild (especially Word) often has mismatched end tags.
        #
        # Default: conservative close only when it matches the current top-of-stack.
        # Exception: for container-ish tags, allow closing even when nested tags
        # (like <figcaption>) are still open, by popping until the requested tag.
        force_close = {
            "html",
            "head",
            "body",
            "figure",
            "table",
            "thead",
            "tbody",
            "tfoot",
            "tr",
            "td",
            "th",
            "ol",
            "ul",
        }
        if self._emit_stack[-1] == tag:
            self._emit_stack.pop()
            self._out.append(f"</{tag}>")
            return

        if tag not in force_close or tag not in self._emit_stack:
            return

        while self._emit_stack:
            top = self._emit_stack.pop()
            self._out.append(f"</{top}>")
            if top == tag:
                break

    def handle_starttag(self, tag: str, attrs: Sequence[Tuple[str, Optional[str]]]) -> None:
        tag = self._map_tag(tag)
        if tag.startswith(WORD_NAMESPACE_PREFIXES):
            return
        if self._skip_depth:
            self._skip_depth += 1
            return
        if self._should_skip_tag(tag):
            self._skip_depth = 1
            self._skip_tag = tag
            return
        self._emit_start(tag, self._filter_attrs(tag, attrs))

    def handle_startendtag(self, tag: str, attrs: Sequence[Tuple[str, Optional[str]]]) -> None:
        tag = self._map_tag(tag)
        if tag.startswith(WORD_NAMESPACE_PREFIXES):
            return
        if self._skip_depth:
            return
        if self._should_skip_tag(tag):
            return
        self._emit_start(tag, self._filter_attrs(tag, attrs))

    def handle_endtag(self, tag: str) -> None:
        tag = self._map_tag(tag)
        if tag.startswith(WORD_NAMESPACE_PREFIXES):
            return
        if self._skip_depth:
            self._skip_depth -= 1
            if self._skip_depth == 0:
                self._skip_tag = None
            return
        self._emit_end(tag)

    def handle_data(self, data: str) -> None:
        if self._skip_depth:
            return
        if not data:
            return
        self._out.append(html.escape(data, quote=False))


@dataclass(frozen=True)
class CleanupOptions:
    fix_images: bool
    merge_tables: bool
    numbered_headings: bool
    keep_img_dimensions: bool
    png_preferred: bool
    img_prefix: str
    light_before_chapter: Optional[int]


def _parse_xml(html_text: str):
    import xml.etree.ElementTree as ET

    try:
        return ET.fromstring(html_text)
    except ET.ParseError:
        # Some Word/Office HTML exports include trailing junk after </html>.
        # Wrap into a container element and take the first <html> if present.
        container = ET.fromstring(f"<root>{html_text}</root>")
        html_elems = list(container.findall("./html"))
        if not html_elems:
            html_elems = list(container.findall(".//html"))
        if not html_elems:
            return container

        base = html_elems[0]
        body = base.find("./body")
        if body is None:
            body = ET.SubElement(base, "body")

        # Merge subsequent <html> fragments into the base body so we don't lose
        # content that appears after an early </html>.
        for extra in html_elems[1:]:
            extra_body = extra.find("./body")
            if extra_body is not None:
                for ch in list(extra_body):
                    extra_body.remove(ch)
                    body.append(ch)
            else:
                for ch in list(extra):
                    extra.remove(ch)
                    body.append(ch)

        # Also keep any non-<html> top-level elements in the wrapped container.
        for ch in list(container):
            if ch.tag == "html":
                continue
            container.remove(ch)
            body.append(ch)

        return base


def _serialize_xml(root) -> str:
    import xml.etree.ElementTree as ET

    # Ensure UTF-8 output without XML declaration.
    return ET.tostring(root, encoding="unicode", method="html")


def _iter_children(parent) -> List:
    return list(parent)


def _is_ignorable_between_tables(node) -> bool:
    tag = (node.tag or "").lower()
    if tag in {"br", "hr"}:
        return True
    if tag == "p":
        text = _norm_text("".join(node.itertext()))
        return text == ""
    # Allow explicit page-break divs.
    if tag == "div":
        cls = (node.attrib.get("class") or "").lower()
        if "page-break" in cls:
            return True
        text = _norm_text("".join(node.itertext()))
        return text == ""
    return False


def _first_row_cells(table) -> Optional[List[str]]:
    # Find the first tr with at least one cell.
    for tr in table.findall(".//tr"):
        cells = tr.findall("./th") + tr.findall("./td")
        if not cells:
            continue
        return [_norm_text("".join(c.itertext())) for c in cells]
    return None


def _effective_col_count(table) -> Optional[int]:
    for tr in table.findall(".//tr"):
        cells = tr.findall("./th") + tr.findall("./td")
        if not cells:
            continue
        count = 0
        for c in cells:
            colspan = c.attrib.get("colspan")
            if colspan and colspan.isdigit():
                count += int(colspan)
            else:
                count += 1
        return count
    return None


def _merge_tables_in_parent(parent) -> bool:
    changed = False
    children = _iter_children(parent)
    i = 0
    while i < len(children):
        node = children[i]
        if (node.tag or "").lower() != "table":
            i += 1
            continue

        j = i + 1
        while j < len(children) and _is_ignorable_between_tables(children[j]):
            j += 1

        if j >= len(children):
            break

        nxt = children[j]
        if (nxt.tag or "").lower() != "table":
            i = j
            continue

        cols_a = _effective_col_count(node)
        cols_b = _effective_col_count(nxt)
        if not cols_a or not cols_b or cols_a != cols_b:
            i = j
            continue

        # Optionally drop a repeated header row on the second table.
        row_a = _first_row_cells(node) or []
        row_b = _first_row_cells(nxt) or []
        drop_first_row_b = bool(row_a and row_b and row_a == row_b)

        # Locate/ensure tbodies
        tbody_a = node.find("./tbody") or node
        tbody_b = nxt.find("./tbody") or nxt
        rows_b = list(tbody_b.findall("./tr"))
        if drop_first_row_b and rows_b:
            rows_b = rows_b[1:]

        for tr in rows_b:
            tbody_a.append(tr)

        # Remove separators and second table
        for k in range(j, i, -1):
            # remove from parent: children[k] .. children[i+1]
            parent.remove(children[k])
        changed = True
        # Refresh view after mutation
        children = _iter_children(parent)
        # Keep `i` to allow chaining merges across multiple split tables.
    return changed


def merge_adjacent_tables(root) -> bool:
    changed = False
    # Only merge tables at the same parent level (common in Docling/Word output).
    for parent in root.findall(".//*"):
        if _merge_tables_in_parent(parent):
            changed = True
    return changed


def _is_ignorable_between_lists(node) -> bool:
    tag = (node.tag or "").lower()
    if tag in {"br", "hr"}:
        return True
    if tag in {"p", "div"}:
        text = _norm_text("".join(node.itertext()))
        if text == "":
            return True
        # Page number / roman numeral artifacts from page breaks.
        if _RE_PAGE_NUM.match(text):
            return True
        cls = (node.attrib.get("class") or "").lower()
        if "page-break" in cls:
            return True
    return False


def _merge_lists_in_parent(parent) -> bool:
    changed = False
    children = _iter_children(parent)
    i = 0
    while i < len(children):
        node = children[i]
        tag = (node.tag or "").lower()
        if tag not in {"ol", "ul"}:
            i += 1
            continue

        j = i + 1
        while j < len(children) and _is_ignorable_between_lists(children[j]):
            j += 1
        if j >= len(children):
            break

        nxt = children[j]
        nxt_tag = (nxt.tag or "").lower()
        if nxt_tag != tag:
            i = j
            continue

        # Merge: append all li from nxt into node
        for li in list(nxt.findall("./li")):
            node.append(li)

        # Remove separators and nxt list
        for k in range(j, i, -1):
            parent.remove(children[k])
        changed = True
        children = _iter_children(parent)
        # Keep `i` to allow chaining merges across multiple split lists.
    return changed


def merge_adjacent_lists(root) -> bool:
    changed = False
    for parent in root.findall(".//*"):
        if _merge_lists_in_parent(parent):
            changed = True
    return changed


def _strip_empty_formatting(root) -> bool:
    # Remove empty <strong>/<em> tags left behind by some Word exports.
    import xml.etree.ElementTree as ET

    changed = False
    parent_map = {c: p for p in root.iter() for c in p}
    for el in list(root.iter()):
        tag = (el.tag or "").lower()
        if tag not in {"strong", "em"}:
            continue
        text = _norm_text(el.text or "")
        if text:
            continue
        if list(el):
            continue
        parent = parent_map.get(el)
        if parent is None:
            continue
        parent.remove(el)
        changed = True
    return changed


def normalize_word_footnotes(root) -> bool:
    """
    Normalize Word-style footnotes into:
      - superscript anchor refs in-text (preserved)
      - a single ordered list of footnotes at the end of <body>

    This is intentionally conservative:
    - Only hoists footnote definition paragraphs that are *direct children* of <body>
      and contain an <a name/id="_ftnN" ...> (or endnote variants).
    - Leaves any nested definitions (e.g., inside tables) untouched.
    """
    import xml.etree.ElementTree as ET

    body = root.find(".//body")
    if body is None:
        body = root

    changed = False

    def _ensure_anchor_ids(a: ET.Element, key: str) -> None:
        nonlocal changed
        if not a.attrib.get("id"):
            a.attrib["id"] = key
            changed = True
        if not a.attrib.get("name"):
            a.attrib["name"] = key
            changed = True

    def _parse_href_num(href: str) -> Optional[Tuple[str, str]]:
        # Returns ("ftn"|"edn", "N") for #_ftnN, #_ednN, #_ftnrefN, #_ednrefN
        href = href.strip()
        if not href.startswith("#_"):
            return None
        frag = href[1:]  # keep leading _
        m = _RE_FTN_REF.match(frag)
        if m:
            # frag looks like _ftnrefN / _ednrefN
            kind = "edn" if frag.lower().startswith("_edn") else "ftn"
            return kind, m.group(1)
        m = _RE_FTN_DEF.match(frag)
        if m:
            kind = "edn" if frag.lower().startswith("_edn") else "ftn"
            return kind, m.group(1)
        return None

    def _make_sup(num: str) -> ET.Element:
        sup = ET.Element("sup")
        sup.text = num
        return sup

    # Normalize in-text references:
    #   <a href="#_ftn1">[1]</a>  ->  <a href="#_ftn1" id/name="_ftnref1"><sup>1</sup></a>
    for a in root.findall(".//a"):
        href = (a.attrib.get("href") or "").strip()
        parsed = _parse_href_num(href)
        if not parsed:
            continue
        kind, num = parsed
        # Only treat href="#_ftnN"/"#_ednN" as references (not the backlinks).
        if href.lower().startswith(f"#_{kind}ref"):
            continue
        ref_key = f"_{kind}ref{num}"
        _ensure_anchor_ids(a, ref_key)

        # If the anchor has no <sup>, and its visible text looks like [N] or N, wrap it.
        has_sup = any((c.tag or "").lower() == "sup" for c in list(a))
        if not has_sup:
            visible = _norm_text("".join(a.itertext()))
            if visible in {num, f"[{num}]"}:
                saved_tail = a.tail
                a.clear()
                a.tag = "a"
                a.attrib["href"] = f"#_{kind}{num}"
                a.attrib["id"] = ref_key
                a.attrib["name"] = ref_key
                a.append(_make_sup(num))
                a.tail = saved_tail
                changed = True

    defs: List[Tuple[str, Optional[str], ET.Element]] = []
    for ch in list(body):
        if (ch.tag or "").lower() != "p":
            continue
        anchors = list(ch.findall(".//a"))
        if not anchors:
            continue

        def_id: Optional[str] = None
        back_href: Optional[str] = None

        # Pattern A: Word-style definition with explicit name/id.
        for a in anchors:
            key = (a.attrib.get("name") or a.attrib.get("id") or "").strip()
            if not key:
                continue
            if _RE_FTN_DEF.match(key):
                def_id = key
                back_href = (a.attrib.get("href") or "").strip() or None
                _ensure_anchor_ids(a, key)
                break

        # Pattern B: Washed definition like <p><a href="#_ftnrefN">[N]</a> ...</p>
        if def_id is None:
            first_a = anchors[0]
            href = (first_a.attrib.get("href") or "").strip()
            parsed = _parse_href_num(href)
            if parsed and href.lower().startswith(f"#_{parsed[0]}ref"):
                kind, num = parsed
                def_id = f"_{kind}{num}"
                back_href = f"#_{kind}ref{num}"

        if not def_id:
            continue

        defs.append((def_id, back_href, ch))
        body.remove(ch)
        changed = True

    if not defs:
        return changed

    # Append footnotes list at end of body.
    body.append(ET.Element("hr"))
    ol = ET.Element("ol")
    for def_id, back_href, p in defs:
        li = ET.Element("li")

        # Add an explicit target anchor so #_ftnN always lands correctly.
        target = ET.SubElement(li, "a")
        target.attrib["id"] = def_id
        target.attrib["name"] = def_id
        target.text = ""

        # Move paragraph contents into the list item, dropping the leading marker anchor if present.
        # Common Word pattern:
        #   <p><a name="_ftn1" href="#_ftnref1">[1]</a> Footnote text...</p>
        if p.text:
            # Usually empty, but keep if present.
            target.tail = (target.tail or "") + p.text
            p.text = None

        for node in list(p):
            if (node.tag or "").lower() == "a":
                node_key = (node.attrib.get("name") or node.attrib.get("id") or "").strip()
                node_href = (node.attrib.get("href") or "").strip()
                # Drop the leading marker anchor:
                # - Word-style: <a name/id="_ftnN" ...>[N]</a>
                # - Washed-style: <a href="#_ftnrefN">[N]</a>
                if node_key.lower() == def_id.lower() or (back_href and node_href == back_href):
                    if node.tail:
                        target.tail = (target.tail or "") + node.tail
                    p.remove(node)
                    continue
            p.remove(node)
            li.append(node)

        # Optional backlink (only if Word provided one).
        if back_href and back_href.startswith("#"):
            back = ET.SubElement(li, "a")
            back.attrib["href"] = back_href
            back.text = "↩"

        ol.append(li)

    body.append(ol)
    return True


def _find_chapter_boundary(body, chapter_num: int) -> Optional[int]:
    """
    Returns the index (in body direct children) where "chapter N" starts.

    Heuristics:
    - Match headings/paragraphs starting with "N." but not "N.x"
    - Also match "Chapitre N" / "CHAPITRE N"
    - Ignore references like "Article N" or "6.2" etc.
    """
    # Examples to match:
    #   "6. Atténuation ..."
    #   "CHAPITRE 6 ..."
    re_chap = re.compile(rf"^\s*chapitre\s*{chapter_num}\b", re.IGNORECASE)
    re_num = re.compile(rf"^\s*{chapter_num}\.(?!\d)\s*\S", re.IGNORECASE)
    re_article = re.compile(rf"^\s*article\s*{chapter_num}\b", re.IGNORECASE)

    for idx, el in enumerate(list(body)):
        tag = (el.tag or "").lower()
        if tag not in HEADING_TAGS and tag not in {"p", "div"}:
            continue
        text = _norm_text("".join(el.itertext()))
        if not text:
            continue
        if re_article.match(text):
            continue
        if re_chap.match(text) or re_num.match(text):
            return idx
    return None


def demote_headings_in_tables(root) -> bool:
    changed = False
    for table in root.findall(".//table"):
        for tag in HEADING_TAGS:
            for h in table.findall(f".//{tag}"):
                text = _norm_text("".join(h.itertext()))
                h.clear()
                h.tag = "p"
                h.text = text
                changed = True
    return changed


def demote_table_figure_titles(root) -> bool:
    # Convert headings that are really captions into <p><strong>...</strong></p>.
    changed = False
    for el in list(root.iter()):
        tag = (el.tag or "").lower()
        if tag not in HEADING_TAGS:
            continue
        text = _norm_text("".join(el.itertext()))
        if not text:
            continue
        low = text.lower()
        if not (low.startswith("tableau") or low.startswith("figure")):
            continue
        el.clear()
        el.tag = "p"
        import xml.etree.ElementTree as ET

        strong = ET.SubElement(el, "strong")
        strong.text = text
        changed = True
    return changed


def remove_empty_headings(root) -> bool:
    import xml.etree.ElementTree as ET

    changed = False
    parent_map = {c: p for p in root.iter() for c in p}
    for el in list(root.iter()):
        tag = (el.tag or "").lower()
        if tag not in HEADING_TAGS:
            continue
        text = _norm_text("".join(el.itertext()))
        if text:
            continue
        if list(el):
            continue
        parent = parent_map.get(el)
        if parent is None:
            continue
        parent.remove(el)
        changed = True
    return changed


def demote_colon_headings(root) -> bool:
    """
    Demote headings like "Résultats escomptés :" that should not be section-level
    headers. Heuristic: heading text ends with ":" and is not a numbered title.
    """
    import xml.etree.ElementTree as ET

    changed = False
    for el in list(root.iter()):
        tag = (el.tag or "").lower()
        if tag not in {"h2", "h3", "h4", "h5", "h6"}:
            continue
        text = _norm_text("".join(el.itertext()))
        if not text:
            continue
        if not text.endswith(":"):
            continue
        if _RE_NUMBERED.match(text):
            continue

        el.clear()
        el.tag = "p"
        strong = ET.SubElement(el, "strong")
        strong.text = text
        changed = True

    return changed


def unwrap_runaway_figures(root) -> bool:
    """
    Fix a common Word/HTML-malformation pattern where a <figure><figcaption> starts
    but never properly closes until the end of the document, causing huge chunks
    of content to render "inside" a figcaption.

    Heuristic: if a <figure> is a direct child of <body> and contains substantial
    document structure (nested figures, headings, lists, tables), unwrap it.
    """
    changed = False
    body = root.find(".//body")
    if body is None:
        return False

    body_children = list(body)

    def _splice(parent, idx: int, nodes: List) -> None:
        for offset, n in enumerate(nodes):
            parent.insert(idx + offset, n)

    for i, child in enumerate(body_children):
        if (child.tag or "").lower() != "figure":
            continue

        # Only unwrap runaway figures at the top level.
        fig = child
        has_nested_figure = fig.find(".//figure") is not None
        has_headings = any(fig.find(f".//{t}") is not None for t in ("h1", "h2", "h3"))
        has_tables = fig.find(".//table") is not None
        has_lists = fig.find(".//ol") is not None or fig.find(".//ul") is not None

        # If the figure contains other figure blocks or major structure,
        # it's almost certainly a runaway wrapper.
        if not (has_nested_figure or has_headings or has_tables or has_lists):
            continue

        figcaption = fig.find("./figcaption")
        if figcaption is not None and list(figcaption):
            nodes = list(figcaption)
            for n in nodes:
                figcaption.remove(n)
        else:
            nodes = list(fig)
            for n in nodes:
                fig.remove(n)

        body.remove(fig)
        _splice(body, i, nodes)
        changed = True
        break

    return changed


def unwrap_strong_block_runs(root) -> bool:
    """
    If <strong> wraps block-level elements, it's usually an artifact of broken
    Word HTML (missing </strong>). Unwrap it to avoid huge bold sections.
    """
    import xml.etree.ElementTree as ET

    blockish = {
        "p",
        "div",
        "table",
        "ul",
        "ol",
        "li",
        "figure",
        "figcaption",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
    }

    changed = False
    parent_map = {c: p for p in root.iter() for c in p}

    for strong in list(root.findall(".//strong")):
        if not any((c.tag or "").lower() in blockish for c in list(strong)):
            continue
        parent = parent_map.get(strong)
        if parent is None:
            continue
        if strong not in list(parent):
            continue
        idx = list(parent).index(strong)
        # Preserve leading text as a separate <p> if needed.
        lead = _norm_text(strong.text or "")
        nodes: List = []
        if lead:
            nodes.append(_make_p_with_text(lead, bold=False))
        for ch in list(strong):
            strong.remove(ch)
            nodes.append(ch)
        parent.remove(strong)
        for off, n in enumerate(nodes):
            parent.insert(idx + off, n)
        changed = True

    return changed


def unwrap_paragraphs_with_blocks(root) -> bool:
    """
    If a <p> contains block-level children (often produced by Word/ET parsing),
    unwrap it by lifting its contents to the parent level.
    """
    import xml.etree.ElementTree as ET

    blockish = {
        "p",
        "div",
        "table",
        "ul",
        "ol",
        "li",
        "figure",
        "figcaption",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
    }

    changed = False
    parent_map = {c: p for p in root.iter() for c in p}

    for p in list(root.findall(".//p")):
        if not any((c.tag or "").lower() in blockish for c in list(p)):
            continue
        parent = parent_map.get(p)
        if parent is None:
            continue
        if p not in list(parent):
            continue
        idx = list(parent).index(p)

        nodes: List = []
        lead = _norm_text(p.text or "")
        if lead:
            nodes.append(_make_p_with_text(lead))

        for ch in list(p):
            tail = _norm_text(ch.tail or "")
            ch.tail = None
            p.remove(ch)
            nodes.append(ch)
            if tail:
                nodes.append(_make_p_with_text(tail))

        parent.remove(p)
        for off, n in enumerate(nodes):
            parent.insert(idx + off, n)
        changed = True

    return changed


def _make_p_with_text(text: str, *, bold: bool = False):
    import xml.etree.ElementTree as ET

    p = ET.Element("p")
    if bold:
        strong = ET.SubElement(p, "strong")
        strong.text = text
    else:
        p.text = text
    return p


def normalize_figures(root) -> bool:
    """
    Normalize Word-exported <figure>/<figcaption> blocks.

    Word sometimes emits:
      <p><figure><figcaption>...lots of normal paragraphs...</figcaption></figure>
    which is invalid and causes browsers to render huge chunks as figcaption.

    Strategy:
    - If a <figure> is nested inside a <p>, replace the *whole <p>* with the
      figcaption contents (and/or figcaption text as <p>).
    - If a <figure> is elsewhere, replace just the <figure> with its figcaption contents.
    - If the figcaption is just a label starting with "Figure"/"Tableau", emit it as bold text.
    """
    import xml.etree.ElementTree as ET

    changed = False

    def _figcaption_to_nodes(figcaption) -> List:
        nodes: List = []
        lead_text = _norm_text(figcaption.text or "")
        if not lead_text:
            if not list(figcaption):
                lead_text = _norm_text("".join(figcaption.itertext()))
        if lead_text:
            low = lead_text.lower()
            is_caption_label = low.startswith("figure") or low.startswith("tableau")
            nodes.append(_make_p_with_text(lead_text, bold=is_caption_label))
        for ch in list(figcaption):
            tail = _norm_text(ch.tail or "")
            ch.tail = None
            figcaption.remove(ch)
            nodes.append(ch)
            if tail:
                nodes.append(_make_p_with_text(tail))
        tail_fc = _norm_text(figcaption.tail or "")
        figcaption.tail = None
        if tail_fc:
            nodes.append(_make_p_with_text(tail_fc))
        return nodes

    # Iterate with restarts to avoid stale parent maps during mutation.
    while True:
        mutated = False
        parent_map = {c: p for p in root.iter() for c in p}
        figures = list(root.findall(".//figure"))

        for fig in figures:
            parent = parent_map.get(fig)
            if parent is None:
                continue

            # Build replacement nodes from the whole <figure> so we don't drop
            # images or other content that might be siblings of <figcaption>.
            repl: List = []
            fig_tail = _norm_text(fig.tail or "")
            fig.tail = None
            for ch in list(fig):
                tail = _norm_text(ch.tail or "")
                ch.tail = None
                fig.remove(ch)
                if (ch.tag or "").lower() == "figcaption":
                    repl.extend(_figcaption_to_nodes(ch))
                else:
                    repl.append(ch)
                if tail:
                    repl.append(_make_p_with_text(tail))
            if fig_tail:
                repl.append(_make_p_with_text(fig_tail))

            if not repl:
                parent.remove(fig)
                changed = True
                mutated = True
                break

            if fig not in list(parent):
                continue
            idx = list(parent).index(fig)
            parent.remove(fig)
            for off, node in enumerate(repl):
                parent.insert(idx + off, node)
            changed = True
            mutated = True
            break

        if not mutated:
            break

    return changed


def apply_numbered_headings(root) -> bool:
    import xml.etree.ElementTree as ET

    changed = False
    parent_map = {c: p for p in root.iter() for c in p}
    for el in root.findall(".//*"):
        tag = (el.tag or "").lower()
        if tag not in {"p", "div"} and tag not in HEADING_TAGS:
            continue

        # Skip if inside a list item or a table cell; we don't want to turn
        # bibliography entries or table content (e.g., years) into headings.
        cur = el
        skip = False
        while cur is not None:
            t = (cur.tag or "").lower()
            if t in {"li", "td", "th", "table"}:
                skip = True
                break
            cur = parent_map.get(cur)
        if skip:
            continue

        text = _norm_text("".join(el.itertext()))
        if not text:
            continue

        m = _RE_NUMBERED.match(text)
        if not m:
            continue

        number = m.group(1)
        dot = m.group(2) or ""
        remainder = m.group(3)
        # Guardrails: avoid decimals/units like "1.5°C" (remainder starts with digit),
        # or unit symbols.
        if remainder[:1].isdigit() or remainder.startswith(("°", "%")):
            continue

        depth = number.count(".") + 1
        # For single-level headings, require an explicit trailing dot (e.g. "1. Title").
        # This avoids promoting year-like lines such as "2026 - 2035".
        if depth == 1 and not dot:
            continue

        new_tag = "h2" if depth == 1 else "h3" if depth == 2 else "h4"
        el.tag = new_tag
        changed = True
    return changed


def _load_image_to_png(src_path: pathlib.Path, dest_png: pathlib.Path) -> None:
    from PIL import Image

    dest_png.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(src_path) as im:
        # Ensure compatibility and deterministic output.
        if im.mode in {"RGBA", "LA"}:
            im = im.convert("RGBA")
        else:
            im = im.convert("RGB")
        im.save(dest_png, format="PNG", optimize=True)


def fix_images(root, *, html_path: pathlib.Path, keep_dimensions: bool, png_preferred: bool) -> Tuple[bool, List[str]]:
    changed = False
    warnings: List[str] = []

    img_dir = html_path.parent / "img"
    img_dir.mkdir(parents=True, exist_ok=True)
    prefix = _safe_stem(html_path.stem)

    mapping: Dict[str, str] = {}
    next_idx = 1

    for img in root.findall(".//img"):
        src = (img.attrib.get("src") or "").strip()
        if not src:
            continue
        if src.startswith(("http://", "https://", "data:")):
            continue

        # Normalize relative paths; handle file:// if it points to a local path
        resolved: Optional[pathlib.Path] = None
        if src.startswith("file://"):
            try:
                # file:///Users/... or file://C:/...
                p = src[len("file://") :]
                resolved = pathlib.Path(p).expanduser()
            except Exception:
                resolved = None
        else:
            resolved = (html_path.parent / src).resolve()

        if not resolved or not resolved.exists():
            warnings.append(f"Missing image file for src={src!r}")
            continue

        key = str(resolved)
        if key in mapping:
            new_name = mapping[key]
        else:
            new_name = f"{prefix}-{next_idx}.png" if png_preferred else f"{prefix}-{next_idx}{resolved.suffix.lower()}"
            next_idx += 1
            mapping[key] = new_name

            dest = img_dir / new_name
            if dest.exists():
                # Don't overwrite; keep existing.
                pass
            else:
                if png_preferred:
                    _load_image_to_png(resolved, dest)
                else:
                    shutil.copyfile(resolved, dest)

        img.attrib["src"] = f"img/{new_name}"
        changed = True

        if not keep_dimensions:
            img.attrib.pop("width", None)
            img.attrib.pop("height", None)

        # Strip Word-specific extra attrs if any survived.
        for k in list(img.attrib.keys()):
            if k not in ALLOWED_ATTRS.get("img", set()):
                img.attrib.pop(k, None)

    return changed, warnings


def fix_images_with_prefix(
    root,
    *,
    html_path: pathlib.Path,
    prefix: str,
    keep_dimensions: bool,
    png_preferred: bool,
) -> Tuple[bool, List[str]]:
    # Small wrapper to control the prefix without changing the HTML filename.
    old_stem = html_path.stem
    try:
        # Create a fake path-like object by swapping just stem via a sibling path.
        fake = html_path.with_name(f"{prefix}{html_path.suffix}")
        return fix_images(
            root,
            html_path=fake,
            keep_dimensions=keep_dimensions,
            png_preferred=png_preferred,
        )
    finally:
        _ = old_stem


def wash_and_parse(html_text: str):
    washer = Washer()
    washer.feed(html_text)
    washed = washer.get_output()
    washed = _strip_illegal_xml_chars(washed)
    return washed, _parse_xml(washed)


def _strip_illegal_xml_chars(text: str) -> str:
    # XML 1.0 valid chars: https://www.w3.org/TR/xml/#charsets
    # Keep: TAB (0x9), LF (0xA), CR (0xD), and U+0020..U+D7FF, U+E000..U+FFFD.
    out_chars: List[str] = []
    for ch in text:
        code = ord(ch)
        if code in (0x9, 0xA, 0xD) or (0x20 <= code <= 0xD7FF) or (0xE000 <= code <= 0xFFFD):
            out_chars.append(ch)
    return "".join(out_chars)


_TABLE_STYLE = (
    "table{border-collapse:collapse}"
    "td,th{border:1px solid #ccc;padding:4px}"
)


def _inject_table_style(root) -> None:
    import xml.etree.ElementTree as ET

    head = root.find(".//head")
    if head is None:
        head = ET.SubElement(root, "head")
    style = ET.SubElement(head, "style")
    style.text = _TABLE_STYLE


def cleanup_file(path: pathlib.Path, *, options: CleanupOptions, backup: bool) -> List[str]:
    html_text = _read_text_guess_encoding(path)

    washed, root = wash_and_parse(html_text)

    warnings: List[str] = []
    changed = washed != html_text

    if unwrap_runaway_figures(root):
        changed = True

    if normalize_figures(root):
        changed = True

    if unwrap_strong_block_runs(root):
        changed = True

    if unwrap_paragraphs_with_blocks(root):
        changed = True

    body = root.find(".//body")
    if body is None:
        body = root

    def _apply_full_transforms(scope_root) -> None:
        nonlocal changed, warnings
        if demote_headings_in_tables(scope_root):
            changed = True
        if normalize_word_footnotes(scope_root):
            changed = True
        if options.merge_tables:
            if merge_adjacent_tables(scope_root):
                changed = True
        if merge_adjacent_lists(scope_root):
            changed = True
        if options.numbered_headings:
            if apply_numbered_headings(scope_root):
                changed = True
        if demote_table_figure_titles(scope_root):
            changed = True
        if demote_colon_headings(scope_root):
            changed = True
        if remove_empty_headings(scope_root):
            changed = True
        if _strip_empty_formatting(scope_root):
            changed = True

    def _apply_images(scope_root) -> None:
        nonlocal changed, warnings
        if not options.fix_images:
            return
        img_changed, img_warnings = fix_images_with_prefix(
            scope_root,
            html_path=path,
            prefix=options.img_prefix or path.stem,
            keep_dimensions=options.keep_img_dimensions,
            png_preferred=options.png_preferred,
        )
        warnings.extend(img_warnings)
        if img_changed:
            changed = True

    if options.light_before_chapter:
        boundary = _find_chapter_boundary(body, options.light_before_chapter)
        if boundary is None:
            # Fallback to full transforms if we can't find the boundary.
            _apply_full_transforms(root)
            _apply_images(root)
        else:
            children = list(body)
            import xml.etree.ElementTree as ET

            pre = ET.Element("div")
            post = ET.Element("div")
            for idx, ch in enumerate(children):
                body.remove(ch)
                (pre if idx < boundary else post).append(ch)

            # Light: only image normalization (and figure unwrapping already happened globally).
            _apply_images(pre)

            # Full: structural cleanup + images.
            _apply_full_transforms(post)
            _apply_images(post)

            for ch in list(pre):
                pre.remove(ch)
                body.append(ch)
            for ch in list(post):
                post.remove(ch)
                body.append(ch)
            changed = True
    else:
        _apply_full_transforms(root)
        _apply_images(root)

    _inject_table_style(root)
    changed = True

    out_html = _serialize_xml(root)
    if backup:
        backup_path = path.with_suffix(path.suffix + ".bak")
        if not backup_path.exists():
            backup_path.write_text(html_text, encoding="utf-8")
    path.write_text(out_html, encoding="utf-8")
    return warnings


def _read_text_guess_encoding(path: pathlib.Path) -> str:
    data = path.read_bytes()

    # BOM-based detection first (covers Word/Office HTML exports well).
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        return data.decode("utf-16", errors="replace")
    if data.startswith(b"\xff\xfe\x00\x00") or data.startswith(b"\x00\x00\xfe\xff"):
        return data.decode("utf-32", errors="replace")
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig", errors="replace")

    for enc in ("utf-8", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _iter_html_files(paths: Iterable[pathlib.Path]) -> List[pathlib.Path]:
    out: List[pathlib.Path] = []
    for p in paths:
        if p.is_dir():
            out.extend(sorted(p.glob("*.html")))
        else:
            out.append(p)
    return out


def cli() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Wash HTML and normalize images into an ./img/ folder next to the processed HTML file(s)"
    )
    ap.add_argument("paths", nargs="+", help="HTML file(s) or directory(ies) containing HTML")
    ap.add_argument("--no-backup", action="store_true", help="Do not write .bak backups")
    ap.add_argument("--no-images", action="store_true", help="Do not rewrite/copy images into ./img/")
    ap.add_argument("--no-merge-tables", action="store_true", help="Disable adjacent-table merging")
    ap.add_argument("--no-numbered-headings", action="store_true", help="Disable numbered heading promotion")
    ap.add_argument(
        "--strip-img-dimensions",
        action="store_true",
        help="Remove width/height attributes on <img> (default keeps them)",
    )
    ap.add_argument(
        "--keep-original-img-format",
        action="store_true",
        help="Copy images as-is instead of converting to PNG",
    )
    ap.add_argument(
        "--img-prefix",
        default="",
        help="Prefix for images written to ./img/ (default: HTML filename stem)",
    )
    ap.add_argument(
        "--light-before-chapter",
        type=int,
        default=0,
        help="Only normalize images before this chapter number; fully clean from this chapter onward (e.g., 6)",
    )
    return ap.parse_args()


def main() -> None:
    args = cli()
    paths = [pathlib.Path(p).expanduser() for p in args.paths]
    html_files = _iter_html_files(paths)
    if not html_files:
        raise SystemExit("No .html files found.")

    options = CleanupOptions(
        fix_images=not args.no_images,
        merge_tables=not args.no_merge_tables,
        numbered_headings=not args.no_numbered_headings,
        keep_img_dimensions=not args.strip_img_dimensions,
        png_preferred=not args.keep_original_img_format,
        img_prefix=args.img_prefix.strip(),
        light_before_chapter=(args.light_before_chapter or None),
    )

    had_warn = False
    for f in html_files:
        warns = cleanup_file(f, options=options, backup=not args.no_backup)
        if warns:
            had_warn = True
            for w in warns:
                print(f"[WARN] {f}: {w}")
        print(f"[OK] {f}")

    if had_warn:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
