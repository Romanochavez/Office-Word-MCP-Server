"""
Header and footer tools for the Word Document Server.

Covers the gap where the membrete lives in a TABLE inside the header (not a plain
paragraph): reading/writing must walk header/footer *tables* and *text boxes*, and
merge text that Word split across multiple runs (the folio/date in real letters are
split across several <w:r> runs).
"""
import os
import json
from typing import Optional, List

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from word_document_server.utils.file_utils import ensure_docx_extension
from word_document_server.utils.document_utils import paragraph_full_text, _replace_in_paragraph

_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_QN_T = qn("w:t")
_QN_TXBX = qn("w:txbxContent")


def _add_page_field(paragraph) -> None:
    """Append a PAGE field run to a paragraph."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _set_story_text(container, text: str, alignment: str, include_page_number: bool) -> None:
    paragraph = container.paragraphs[0] if container.paragraphs else container.add_paragraph()
    paragraph.text = ""
    run = paragraph.add_run(text or "")
    if include_page_number:
        run.add_text("   ")
        _add_page_field(paragraph)
    if alignment and alignment.lower() in _ALIGN:
        paragraph.alignment = _ALIGN[alignment.lower()]


# ---------------------------------------------------------------------------
# Helpers para leer/escribir TODAS las historias (párrafos + tablas + text boxes)
# ---------------------------------------------------------------------------

def _story_paragraphs(story) -> List:
    """All Paragraph objects in a header/footer story (direct + inside table cells)."""
    paras = list(story.paragraphs)
    for table in story.tables:
        for row in table.rows:
            for cell in row.cells:
                paras.extend(cell.paragraphs)
    return paras


def _story_tables_text(story):
    out = []
    for table in story.tables:
        rows = []
        for row in table.rows:
            cells = []
            for cell in row.cells:
                cells.append("\n".join(paragraph_full_text(p) for p in cell.paragraphs))
            rows.append(cells)
        out.append(rows)
    return out


def _story_full_text(story) -> str:
    parts = [paragraph_full_text(p) for p in story.paragraphs]
    for table in story.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    parts.append(paragraph_full_text(p))
    return " ".join(x for x in parts if x.strip())


def _alternate_twin(tb_elem):
    """Return the sibling txbxContent of the same logical box inside mc:AlternateContent."""
    anc = tb_elem.getparent()
    while anc is not None:
        tag = anc.tag.split("}")[-1]
        if tag in ("Choice", "Fallback"):
            ac = anc.getparent()
            other = None
            for child in ac:
                ctag = child.tag.split("}")[-1]
                if ctag in ("Choice", "Fallback") and child is not anc:
                    other = child
            if other is not None:
                tbs = list(other.iter(_QN_TXBX))
                if tbs:
                    return tbs[0]
            return None
        anc = anc.getparent()
    return None


def _all_textboxes(doc):
    """Stable ordered list of (location, txbxContent element) across body + headers + footers."""
    locs = [("body", doc.element.body)]
    for i, sec in enumerate(doc.sections):
        locs.append(("header%d" % i, sec.header._element))
        locs.append(("footer%d" % i, sec.footer._element))
    out = []
    for loc, el in locs:
        for tb in el.iter(_QN_TXBX):
            out.append((loc, tb))
    return out


def _textbox_text(tb_elem) -> str:
    return "".join((t.text or "") for t in tb_elem.iter(_QN_T))


def _set_textbox_text(tb_elem, new_text: str) -> None:
    ts = list(tb_elem.iter(_QN_T))
    if not ts:
        p = OxmlElement("w:p"); r = OxmlElement("w:r"); t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve"); t.text = new_text
        r.append(t); p.append(r); tb_elem.append(p)
        return
    ts[0].set(qn("xml:space"), "preserve")
    ts[0].text = new_text
    for t in ts[1:]:
        t.text = ""


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

async def add_header(filename: str, text: str = "", alignment: str = "left",
                     different_first_page: bool = False,
                     first_page_text: Optional[str] = None) -> str:
    """Set the page header of the document's first section."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        doc = Document(filename)
        section = doc.sections[0]
        if different_first_page:
            section.different_first_page_header_footer = True
        _set_story_text(section.header, text, alignment, include_page_number=False)
        if different_first_page and first_page_text is not None:
            _set_story_text(section.first_page_header, first_page_text, alignment, include_page_number=False)
        doc.save(filename)
        return json.dumps({"ok": True, "message": f"Header set on {filename}",
                           "different_first_page": different_first_page})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


async def add_footer(filename: str, text: str = "", alignment: str = "center",
                     different_first_page: bool = False,
                     first_page_text: Optional[str] = None,
                     include_page_number: bool = False) -> str:
    """Set the page footer of the document's first section."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        doc = Document(filename)
        section = doc.sections[0]
        if different_first_page:
            section.different_first_page_header_footer = True
        _set_story_text(section.footer, text, alignment, include_page_number)
        if different_first_page and first_page_text is not None:
            _set_story_text(section.first_page_footer, first_page_text, alignment, include_page_number)
        doc.save(filename)
        return json.dumps({"ok": True, "message": f"Footer set on {filename}",
                           "include_page_number": include_page_number})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


async def get_header_footer(filename: str) -> str:
    """Read header and footer content of each section, including TABLES and TEXT BOXES
    with cross-run text merged. Read-only."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"error": f"Document {filename} does not exist"})
    try:
        doc = Document(filename)
        boxes = [(loc, tb) for loc, tb in _all_textboxes(doc)]
        out = []
        for i, section in enumerate(doc.sections):
            hdr_boxes = [{"index": j, "text": _textbox_text(tb)}
                         for j, (loc, tb) in enumerate(boxes) if loc == f"header{i}"]
            ftr_boxes = [{"index": j, "text": _textbox_text(tb)}
                         for j, (loc, tb) in enumerate(boxes) if loc == f"footer{i}"]
            out.append({
                "section": i,
                "header": {
                    "paragraphs": [paragraph_full_text(p) for p in section.header.paragraphs],
                    "tables": _story_tables_text(section.header),
                    "textboxes": hdr_boxes,
                    "fullText": _story_full_text(section.header),
                },
                "footer": {
                    "paragraphs": [paragraph_full_text(p) for p in section.footer.paragraphs],
                    "tables": _story_tables_text(section.footer),
                    "textboxes": ftr_boxes,
                    "fullText": _story_full_text(section.footer),
                },
                "different_first_page": bool(section.different_first_page_header_footer),
            })
        return json.dumps(out, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


async def replace_in_header_footer(filename: str, find_text: str, replace_text: str,
                                   story: str = "both", occurrence: Optional[int] = None) -> str:
    """Find and replace text in headers/footers, including table cells and text boxes.

    Cross-run aware: matches text split across multiple runs (e.g. folio 'MEXL1-MRRC-CA-02879'
    stored as several runs). `story`: 'header', 'footer' or 'both'. `occurrence`: 1-based
    occurrence to replace (None = all).
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    if not find_text:
        return json.dumps({"ok": False, "error": "find_text no puede estar vacío"})
    story = (story or "both").lower()
    try:
        doc = Document(filename)
        stories = []
        if story in ("header", "both"):
            stories.extend((f"header{i}", s.header) for i, s in enumerate(doc.sections))
        if story in ("footer", "both"):
            stories.extend((f"footer{i}", s.footer) for i, s in enumerate(doc.sections))

        total = 0
        seen = {}
        for name, st in stories:
            for para in _story_paragraphs(st):
                c, _ = _replace_in_paragraph(para, find_text, replace_text)
                total += c
            for tb in st._element.iter(_QN_TXBX):
                text = _textbox_text(tb)
                if find_text in text:
                    if occurrence is None or (seen.get(name, 0) < occurrence):
                        new_text = text.replace(find_text, replace_text, 1 if occurrence else -1)
                        _set_textbox_text(tb, new_text)
                        twin = _alternate_twin(tb)
                        if twin is not None:
                            _set_textbox_text(twin, new_text)
                        seen[name] = seen.get(name, 0) + 1
                        total += 1

        doc.save(filename)
        return json.dumps({"ok": True, "replacements": total, "story": story})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


async def get_textboxes(filename: str) -> str:
    """List all text boxes (body + headers + footers) with a stable index. Read-only."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"error": f"Document {filename} does not exist"})
    try:
        doc = Document(filename)
        boxes = []
        for j, (loc, tb) in enumerate(_all_textboxes(doc)):
            boxes.append({"index": j, "location": loc, "text": _textbox_text(tb)})
        return json.dumps({"count": len(boxes), "textboxes": boxes}, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


async def set_textbox_text(filename: str, textbox_index: int, new_text: str) -> str:
    """Replace the whole text of a specific text box, identified by its stable index
    (from get_textboxes). Also updates the mc:AlternateContent twin if present."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        doc = Document(filename)
        boxes = _all_textboxes(doc)
        if textbox_index < 0 or textbox_index >= len(boxes):
            return json.dumps({"ok": False, "error": f"textbox_index fuera de rango (0..{len(boxes)-1})"})
        loc, tb = boxes[textbox_index]
        _set_textbox_text(tb, new_text)
        twin = _alternate_twin(tb)
        if twin is not None:
            _set_textbox_text(twin, new_text)
        doc.save(filename)
        return json.dumps({"ok": True, "index": textbox_index, "location": loc})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})
