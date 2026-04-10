"""
Processes a formatted comunicado .docx and produces a plain-formatted version:
- Aptos font, 12pt
- No paragraph spacing (space_before/space_after = 0)
- Single line spacing
- Empty paragraph between each content paragraph
- Justified text
"""

import os
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree


def _get_para_text(para_el):
    """Extract plain text from a <w:p> element."""
    texts = []
    for t in para_el.iter(qn('w:t')):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


def _get_table_rows(tbl_el):
    """Extract rows of text from a <w:tbl> element."""
    rows = []
    for tr in tbl_el.findall('.//' + qn('w:tr')):
        cells = []
        for tc in tr.findall('.//' + qn('w:tc')):
            cell_texts = []
            for t in tc.iter(qn('w:t')):
                if t.text:
                    cell_texts.append(t.text)
            cell_text = ''.join(cell_texts).strip()
            if cell_text:
                cells.append(cell_text)
        if cells:
            rows.append('\t'.join(cells))
    return rows


def _extract_content_in_order(doc):
    """
    Walk the document body in document order and return a list of non-empty
    text strings, preserving paragraph / table order.
    """
    body = doc.element.body
    content = []

    for child in body:
        tag = child.tag
        if tag == qn('w:p'):
            text = _get_para_text(child)
            if text:
                content.append(text)
        elif tag == qn('w:tbl'):
            for row_text in _get_table_rows(child):
                content.append(row_text)
        # sectPr and other elements are ignored

    return content


def _make_plain_paragraph(doc, text):
    """
    Add a paragraph with plain formatting:
      - Aptos font, 12 pt
      - Justified
      - No space before/after
      - Single line spacing (auto 240)
    """
    para = doc.add_paragraph()

    # Paragraph-level formatting
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(12)

    # Set spacing via XML for precision (matches ComPrensa style)
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    # Replace any existing spacing element
    existing = pPr.find(qn('w:spacing'))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(spacing)

    # Add run with Aptos font
    run = para.add_run(text)
    run.font.name = 'Aptos'
    run.font.size = Pt(12)

    # Set the theme font via XML so Word resolves to Aptos (body) correctly
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Aptos')
    rFonts.set(qn('w:hAnsi'), 'Aptos')
    rFonts.set(qn('w:asciiTheme'), 'minorHAnsi')
    rFonts.set(qn('w:hAnsiTheme'), 'minorHAnsi')

    return para


def _make_empty_paragraph(doc):
    """Add an empty separator paragraph with the same plain formatting."""
    return _make_plain_paragraph(doc, '')


def _build_output_filename(input_path):
    """Derive a sensible output filename from the input."""
    base = os.path.splitext(os.path.basename(input_path))[0]
    # Strip the input file's own suffix if it ends with _input
    base = re.sub(r'_input$', '', base)
    return f'ComPrensa_{base}_plain.docx'


def process_comunicado(input_path: str, output_path: str) -> str:
    """
    Read a formatted comunicado .docx and write a plain-formatted version.

    Returns the suggested download filename.
    """
    input_doc = Document(input_path)
    content_blocks = _extract_content_in_order(input_doc)

    if not content_blocks:
        raise ValueError('No text content found in the uploaded document.')

    out_doc = Document()

    # Remove the default empty paragraph that python-docx adds
    for p in out_doc.paragraphs:
        p._element.getparent().remove(p._element)

    # Set default document font via styles
    style = out_doc.styles['Normal']
    style.font.name = 'Aptos'
    style.font.size = Pt(12)

    # Write content with empty paragraph separators
    for i, text in enumerate(content_blocks):
        _make_plain_paragraph(out_doc, text)
        if i < len(content_blocks) - 1:
            _make_empty_paragraph(out_doc)

    out_doc.save(output_path)

    filename = _build_output_filename(input_path)
    return filename
