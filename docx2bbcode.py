#!/usr/bin/env python3
"""
Convert a .docx file to phpBB BBCode, preserving:
  - bold / italic / underline / strikethrough
  - hyperlinks, including both real w:hyperlink elements AND
    Word "field code" hyperlinks ({ HYPERLINK "url" }display text{ })
  - nested bulleted/numbered lists (via numPr, regardless of paragraph style)
  - manual line breaks (w:br) and blank spacer paragraphs

Usage:
    python3 docx_to_bbcode.py input.docx > output.bbcode.txt
"""
import sys
import re
import zipfile
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "r": R_NS}

LIST_FMT_TO_BBCODE = {
    "bullet": None,  # plain [list]
    "decimal": "1",
    "decimalZero": "1",
    "lowerLetter": "a",
    "upperLetter": "A",
    "lowerRoman": "i",
    "upperRoman": "I",
}


def qn(tag):
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def local(el):
    return etree.QName(el).localname


def load_parts(path):
    with zipfile.ZipFile(path) as z:
        document_xml = z.read("word/document.xml")
        try:
            numbering_xml = z.read("word/numbering.xml")
        except KeyError:
            numbering_xml = None
        try:
            rels_xml = z.read("word/_rels/document.xml.rels")
        except KeyError:
            rels_xml = None
    return document_xml, numbering_xml, rels_xml


def build_rels_map(rels_xml):
    rels = {}
    if rels_xml is None:
        return rels
    root = etree.fromstring(rels_xml)
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        rels[rid] = target
    return rels


def build_numbering_map(numbering_xml):
    """Return {(numId, ilvl_str): numFmt_str}."""
    result = {}
    if numbering_xml is None:
        return result
    root = etree.fromstring(numbering_xml)
    abstract = {}
    for an in root.findall(qn("w:abstractNum")):
        aid = an.get(qn("w:abstractNumId"))
        lvls = {}
        for lvl in an.findall(qn("w:lvl")):
            ilvl = lvl.get(qn("w:ilvl"))
            fmt_el = lvl.find(qn("w:numFmt"))
            fmt = fmt_el.get(qn("w:val")) if fmt_el is not None else "bullet"
            lvls[ilvl] = fmt
        abstract[aid] = lvls
    num_to_abstract = {}
    for num in root.findall(qn("w:num")):
        num_id = num.get(qn("w:numId"))
        aid_el = num.find(qn("w:abstractNumId"))
        if aid_el is not None:
            num_to_abstract[num_id] = aid_el.get(qn("w:val"))
    for num_id, aid in num_to_abstract.items():
        for ilvl, fmt in abstract.get(aid, {}).items():
            result[(num_id, ilvl)] = fmt
    return result


def get_num_pr(p):
    """Return (numId, ilvl) as strings, or (None, None) if paragraph is not a list item."""
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None, None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None, None
    num_id_el = numPr.find(qn("w:numId"))
    ilvl_el = numPr.find(qn("w:ilvl"))
    num_id = num_id_el.get(qn("w:val")) if num_id_el is not None else None
    ilvl = (ilvl_el.get(qn("w:val")) if ilvl_el is not None else None) or "0"
    if num_id is None or num_id == "0":
        # numId 0 explicitly means "not actually numbered" in OOXML
        return None, None
    return num_id, ilvl


def get_run_props(r):
    rPr = r.find(qn("w:rPr"))
    bold = italic = underline = strike = False
    if rPr is not None:
        b = rPr.find(qn("w:b"))
        if b is not None and b.get(qn("w:val")) not in ("0", "false", "off"):
            bold = True
        i_ = rPr.find(qn("w:i"))
        if i_ is not None and i_.get(qn("w:val")) not in ("0", "false", "off"):
            italic = True
        u = rPr.find(qn("w:u"))
        if u is not None and u.get(qn("w:val")) not in (None, "none"):
            underline = True
        strike_el = rPr.find(qn("w:strike"))
        if strike_el is not None and strike_el.get(qn("w:val")) not in ("0", "false", "off"):
            strike = True
    return bold, italic, underline, strike


def run_text(r):
    parts = []
    for child in r:
        tag = local(child)
        if tag == "t":
            parts.append(child.text or "")
        elif tag == "br":
            parts.append("\n")
        elif tag == "cr":
            parts.append("\n")
        elif tag == "tab":
            parts.append("\t")
        elif tag == "noBreakHyphen":
            parts.append("‑")
    return "".join(parts)


def run_has_fldchar(r, kind):
    for child in r:
        if local(child) == "fldChar" and child.get(qn("w:fldCharType")) == kind:
            return True
    return False


def run_instrtext(r):
    parts = []
    for child in r:
        if local(child) == "instrText":
            parts.append(child.text or "")
    return "".join(parts)


def parse_hyperlink_field(instr):
    """Parse a Word ' HYPERLINK "url" \\l "anchor" ' field instruction into a
    usable URL, appending the \\l bookmark switch (if present) as a #fragment."""
    m = re.search(r'HYPERLINK\s+"([^"]+)"', instr)
    if not m:
        return None
    url = m.group(1)
    anchor_m = re.search(r'\\l\s+"([^"]+)"', instr)
    if anchor_m:
        url = f"{url}#{anchor_m.group(1)}"
    return url


def process_paragraph_inline(p, rels):
    """Walk paragraph children in document order, returning a list of segment dicts:
    {'text', 'bold', 'italic', 'underline', 'strike', 'url'}.
    Handles both real w:hyperlink elements and Word field-code hyperlinks
    (fldChar begin -> instrText HYPERLINK "url" -> fldChar separate -> display runs -> fldChar end).
    """
    segments = []
    field_state = None  # None | 'instr' | 'display'
    instr_buf = ""
    display_segments = []

    def make_seg(r, url=None):
        bold, italic, underline, strike = get_run_props(r)
        text = run_text(r)
        if text == "":
            return None
        return {
            "text": text,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "strike": strike,
            "url": url,
        }

    for child in p:
        tag = local(child)
        if tag == "hyperlink":
            rid = child.get(qn("r:id"))
            anchor = child.get(qn("w:anchor"))
            if rid:
                url = rels.get(rid)
            elif anchor:
                url = "#" + anchor
            else:
                url = None
            for r in child.findall(qn("w:r")):
                seg = make_seg(r, url=url)
                if seg:
                    segments.append(seg)
            continue

        if tag != "r":
            continue  # ignore bookmarkStart/End, proofErr, etc.

        if run_has_fldchar(child, "begin"):
            field_state = "instr"
            instr_buf = ""
            continue

        if field_state == "instr":
            instr_buf += run_instrtext(child)
            if run_has_fldchar(child, "separate"):
                field_state = "display"
            continue

        if field_state == "display":
            if run_has_fldchar(child, "end"):
                url = parse_hyperlink_field(instr_buf)
                for seg in display_segments:
                    seg["url"] = url
                    segments.append(seg)
                display_segments = []
                field_state = None
                continue
            seg = make_seg(child)
            if seg:
                display_segments.append(seg)
            continue

        # Plain run, no active field
        seg = make_seg(child)
        if seg:
            segments.append(seg)

    return segments


def escape_url(url):
    # phpBB's [url=...] value is terminated by ']'; guard against literal ']' breaking it.
    return url.replace("]", "%5D") if url else url


def bbcode_wrap(text, bold, italic, underline, strike, url):
    out = text
    if strike:
        out = f"[s]{out}[/s]"
    if underline:
        out = f"[u]{out}[/u]"
    if italic:
        out = f"[i]{out}[/i]"
    if bold:
        out = f"[b]{out}[/b]"
    if url:
        out = f"[url={escape_url(url)}]{out}[/url]"
    return out


def segments_to_bbcode(segments):
    if not segments:
        return ""
    merged = []
    for seg in segments:
        key = (seg["bold"], seg["italic"], seg["underline"], seg["strike"], seg["url"])
        if merged and merged[-1][0] == key:
            merged[-1] = (key, merged[-1][1] + seg["text"])
        else:
            merged.append((key, seg["text"]))
    parts = []
    for (bold, italic, underline, strike, url), text in merged:
        parts.append(bbcode_wrap(text, bold, italic, underline, strike, url))
    return "".join(parts)


def list_bbcode_open(fmt):
    val = LIST_FMT_TO_BBCODE.get(fmt)
    return f"[list={val}]" if val else "[list]"


def convert(path):
    document_xml, numbering_xml, rels_xml = load_parts(path)
    rels = build_rels_map(rels_xml)
    numbering_map = build_numbering_map(numbering_xml)

    root = etree.fromstring(document_xml)
    body = root.find(qn("w:body"))

    lines = []
    open_levels = []  # list of numFmt strings, index = depth (0-based)

    def close_all_lists():
        while open_levels:
            open_levels.pop()
            lines.append("[/list]")

    for el in body:
        tag = local(el)
        if tag != "p":
            # sectPr or other non-paragraph body-level content: skip (no visible text)
            continue

        num_id, ilvl = get_num_pr(el)
        segments = process_paragraph_inline(el, rels)
        text = segments_to_bbcode(segments)

        if num_id is not None and ilvl is not None:
            ilvl_int = int(ilvl)
            fmt = numbering_map.get((num_id, ilvl), "bullet")

            # Close levels deeper than needed
            while len(open_levels) - 1 > ilvl_int:
                open_levels.pop()
                lines.append("[/list]")

            # Open levels up to needed depth
            while len(open_levels) - 1 < ilvl_int:
                depth = len(open_levels)
                depth_fmt = numbering_map.get((num_id, str(depth)), "bullet")
                lines.append(list_bbcode_open(depth_fmt))
                open_levels.append(depth_fmt)

            # If format at current depth changed (different list type resumed
            # at the same nesting depth), restart the list level.
            if open_levels[ilvl_int] != fmt:
                lines.append("[/list]")
                lines.append(list_bbcode_open(fmt))
                open_levels[ilvl_int] = fmt

            lines.append(f"[*]{text}")
        else:
            close_all_lists()
            lines.append(text)

    close_all_lists()
    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 docx_to_bbcode.py input.docx", file=sys.stderr)
        sys.exit(1)
    print(convert(sys.argv[1]))
