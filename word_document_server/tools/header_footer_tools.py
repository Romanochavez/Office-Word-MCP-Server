"""
Header and footer tools for the Word Document Server.

Adds page headers and footers (section-level) — a gap the server previously had:
the only existing "header" tools manipulate *headings* (body paragraphs), not the
page header/footer. These use python-docx `section.header` / `section.footer`.
"""
import os
import json
from typing import Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from word_document_server.utils.file_utils import ensure_docx_extension

_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _add_page_field(paragraph) -> None:
    """Append a PAGE field run to a paragraph (shows the current page number)."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _set_story_text(container, text: str, alignment: str, include_page_number: bool) -> None:
    """Write `text` (and optional page number) into a header/footer container."""
    paragraph = container.paragraphs[0] if container.paragraphs else container.add_paragraph()
    paragraph.text = ""  # clear any existing runs (keeps paragraph, drops default)
    run = paragraph.add_run(text or "")
    if include_page_number:
        run.add_text("   ")
        _add_page_field(paragraph)
    if alignment and alignment.lower() in _ALIGN:
        paragraph.alignment = _ALIGN[alignment.lower()]


def _resolve(filename: str):
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return filename, None, f"Document {filename} does not exist"
    return filename, None, None


async def add_header(filename: str, text: str = "", alignment: str = "left",
                     different_first_page: bool = False,
                     first_page_text: Optional[str] = None) -> str:
    """Set the page header (encabezado de página) for the document's first section.

    Args:
        filename: Path to the Word document.
        text: Text for the default (odd-page) header.
        alignment: 'left', 'center', 'right' or 'justify'.
        different_first_page: If True, enables a distinct first-page header.
        first_page_text: Text for the first-page header (requires different_first_page=True).
    """
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
    """Set the page footer (pie de página) for the document's first section.

    Args:
        filename: Path to the Word document.
        text: Text for the default (odd-page) footer.
        alignment: 'left', 'center', 'right' or 'justify'.
        different_first_page: If True, enables a distinct first-page footer.
        first_page_text: Text for the first-page footer (requires different_first_page=True).
        include_page_number: If True, appends an auto PAGE-number field after the text.
    """
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
    """Read the header and footer text of each section (read-only)."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"error": f"Document {filename} does not exist"})
    try:
        doc = Document(filename)
        out = []
        for i, section in enumerate(doc.sections):
            out.append({
                "section": i,
                "header": [p.text for p in section.header.paragraphs],
                "footer": [p.text for p in section.footer.paragraphs],
                "different_first_page": bool(section.different_first_page_header_footer),
            })
        return json.dumps(out, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)})
