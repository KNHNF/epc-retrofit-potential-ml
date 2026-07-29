"""
generate_report.py
Generates the Word document (decimal-numbered sections) for the EPC retrofit prediction coursework.
Run AFTER all notebooks have been executed and model_comparison.csv exists.

Two output modes, both built from the same underlying content:
  full  -> report/25065219_report.docx        (quality-first, word count not constrained)
  safe  -> report/25065219_report_safe.docx    (<=2100 words body, hard cap; title/abstract/
                                                 references/footnotes excluded from the count per
                                                 standard academic convention. Extended methodology
                                                 justification and supporting detail that would
                                                 otherwise bloat the body sits in real Word
                                                 footnotes instead.)

Usage: python src/generate_report.py [full|safe|both]   (default: both)
"""

import os
import re
import sys
import csv
import numpy as np

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.opc.part import XmlPart
from docx.opc.packuri import PackURI

from two_column_layout import (
    add_floating_picture, add_bordered_picture, add_equation,
    add_display_equation, disable_heading_widow_control,
    set_compatibility_mode_full,
)

DATA_DIR    = "data/processed"
FIGURES_DIR = "report/figures"

os.makedirs("report", exist_ok=True)

# ----------------------------------------------------------------------
# Real Word footnotes/endnotes. python-docx has no built-in API for either,
# so this builds the word/footnotes.xml or word/endnotes.xml part by hand
# (separator + continuation separator entries are required by the OOXML
# schema even with zero notes). Note text lives in a separate document
# part, so it is NOT picked up by doc.paragraphs and does not count toward
# the body word count below, same convention as excluding references.
#
# Two-column mode uses ENDNOTES, not footnotes. Confirmed in real Word
# (2026-07-23): Word ties a footnote area's column count to the section's
# body column count with no override, so a single footnote long enough to
# spill from column 1's footnote area into column 2's forces Word to end
# the page early, stranding blank space below the shorter column, the same
# failure mode as the earlier section-break and float-overlap bugs, just a
# different mechanism. Endnotes collect once at the very end of the
# document instead of per-page, so they never interact with the section's
# column layout at all, structurally avoiding the whole bug category rather
# than patching around it. Single-column mode keeps footnotes (confirmed
# to render correctly there, no column to spill across).
# ----------------------------------------------------------------------
FOOTNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
FOOTNOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
ENDNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
ENDNOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"

FOOTNOTES_XML_TEMPLATE = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
</w:footnotes>'''

ENDNOTES_XML_TEMPLATE = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
</w:endnotes>'''


class FootnoteManager:
    """Creates the footnotes or endnotes part once per document and appends
    numbered notes, each with a superscript marker in the body paragraph.
    Pass `use_endnotes=True` for two-column mode (see module note above for
    why); the `.add(paragraph, text)` call site stays identical either way."""

    def __init__(self, doc, use_endnotes=False):
        self.kind = 'endnote' if use_endnotes else 'footnote'
        styles = doc.styles
        text_style_name = 'Endnote Text' if use_endnotes else 'Footnote Text'
        ref_style_name = 'Endnote Reference' if use_endnotes else 'Footnote Reference'
        self.ft_style_id = self._ensure_style(styles, text_style_name, WD_STYLE_TYPE.PARAGRAPH, size=Pt(10))
        self.fr_style_id = self._ensure_style(styles, ref_style_name, WD_STYLE_TYPE.CHARACTER, superscript=True)

        if use_endnotes:
            self.element = parse_xml(ENDNOTES_XML_TEMPLATE)
            part = XmlPart(PackURI('/word/endnotes.xml'), ENDNOTES_CT, self.element, doc.part.package)
            doc.part.relate_to(part, ENDNOTES_REL_TYPE)
        else:
            self.element = parse_xml(FOOTNOTES_XML_TEMPLATE)
            part = XmlPart(PackURI('/word/footnotes.xml'), FOOTNOTES_CT, self.element, doc.part.package)
            doc.part.relate_to(part, FOOTNOTES_REL_TYPE)
        self._next_id = 1

    @staticmethod
    def _ensure_style(styles, name, style_type, size=None, superscript=None):
        try:
            style = styles[name]
        except KeyError:
            style = styles.add_style(name, style_type)
            if size is not None:
                style.font.size = size
                style.font.name = 'Times New Roman'
            if superscript is not None:
                style.font.superscript = superscript
        return style.style_id

    def add(self, paragraph, text):
        note_id = self._next_id
        self._next_id += 1
        k = self.kind

        fnote = OxmlElement(f'w:{k}')
        fnote.set(qn('w:id'), str(note_id))
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        pStyle = OxmlElement('w:pStyle')
        pStyle.set(qn('w:val'), self.ft_style_id)
        pPr.append(pStyle)
        p.append(pPr)
        r1 = OxmlElement('w:r')
        rPr1 = OxmlElement('w:rPr')
        rStyle1 = OxmlElement('w:rStyle')
        rStyle1.set(qn('w:val'), self.fr_style_id)
        rPr1.append(rStyle1)
        r1.append(rPr1)
        r1.append(OxmlElement(f'w:{k}Ref'))
        p.append(r1)
        r2 = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = ' ' + text
        r2.append(t)
        p.append(r2)
        fnote.append(p)
        self.element.append(fnote)

        run = paragraph.add_run()
        rPr = OxmlElement('w:rPr')
        rStyle = OxmlElement('w:rStyle')
        rStyle.set(qn('w:val'), self.fr_style_id)
        rPr.append(rStyle)
        run._r.append(rPr)
        ref = OxmlElement(f'w:{k}Reference')
        ref.set(qn('w:id'), str(note_id))
        run._r.append(ref)
        return note_id


# ----------------------------------------------------------------------
# Two-column (IEEE-style) layout support. Body text runs in two columns;
# figures and tables break out to a single full-width column for that one
# section (standard practice, a 6-column results table or a wide heatmap
# is unreadable squeezed into a ~3-inch column) then the layout returns to
# two columns for the next paragraph. Section breaks in OOXML store a
# section's properties in the paragraph that CLOSES it, not the one that
# opens it, so switch_columns() always describes the section ending there.
# ----------------------------------------------------------------------
def set_section_columns(section, num_cols, space_twips=720):
    sectPr = section._sectPr
    cols = sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        sectPr.append(cols)
    cols.set(qn('w:num'), str(num_cols))
    cols.set(qn('w:space'), str(space_twips))


def switch_columns(doc, num_cols):
    section = doc.add_section(WD_SECTION_START.CONTINUOUS)
    set_section_columns(section, num_cols)
    return section


def add_figure(doc, path, label, width=5.5, two_column=False, dense=True):
    """Add a figure inside a bordered box with a short label ("Fig. 1.")
    at the top-left, the technique used by the 92/100 NLP exemplar. What
    the figure actually shows and why it matters belongs in the
    surrounding body prose calling this, `label` is a tag, not a
    description, do not pass a full sentence here.

    Single-column mode: always a real 1x1 bordered table (`add_bordered_
    picture`), a single column already spans the full usable page width,
    so no figure ever needs to cross a column gutter here.

    Two-column mode: no section break, ever, that's what caused the
    original whitespace problem. `dense=True` (multi-panel plots, long
    axis labels, correlation matrices, anything that would go illegible
    shrunk to column width) floats the figure across both columns via
    `add_floating_picture` (a real 1x1 table cannot cross the column
    gutter, confirmed broken three ways, so this uses a floating picture
    with the label as a plain paragraph placed above it instead).
    `dense=False` (a simple single-series chart with short labels) uses
    the same real bordered table as single-column mode, at column width.
    """
    if not os.path.exists(path):
        p = doc.add_paragraph(f"[FIGURE: {label}. Run notebooks to generate.]")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return
    if two_column:
        if dense:
            add_floating_picture(doc, path, label, width_in=6.2)
        else:
            add_bordered_picture(doc, path, label, width_in=2.8)
        return
    add_bordered_picture(doc, path, label, width_in=width)


def load_comparison_csv():
    path = f"{DATA_DIR}/model_comparison.csv"
    if not os.path.exists(path):
        return None
    rows = []
    with open(path, newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def load_naive_baseline():
    path = f"{DATA_DIR}/naive_baseline_metrics.csv"
    if not os.path.exists(path):
        return None
    with open(path, newline='') as f:
        reader = csv.DictReader(f)
        row = next(reader, None)
    if row is None:
        return None
    return {k: (v if k == 'model' else float(v)) for k, v in row.items()}


def load_bristol_case():
    path = f"{DATA_DIR}/bristol_case_study.csv"
    summary_path = f"{DATA_DIR}/bristol_case_study_summary.csv"
    if not (os.path.exists(path) and os.path.exists(summary_path)):
        return None, None
    rows = []
    with open(path, newline='') as f:
        for row in csv.DictReader(f):
            rows.append(row)
    with open(summary_path, newline='') as f:
        summary = next(csv.DictReader(f))
    int_fields = {'n_bristol_test_properties', 'n_rest_of_test'}
    summary = {k: (int(float(v)) if k in int_fields else float(v))
               for k, v in summary.items()}
    return rows, summary


def add_bristol_table(doc, rows):
    display_cols = [
        ('OUTWARD_POSTCODE', 'Postcode'), ('PROPERTY_TYPE', 'Type'),
        ('CURRENT_ENERGY_RATING', 'Rating'), ('EFFICIENCY_GAP', 'Gap (pts)'),
        ('PRED_PROBA', 'P(headroom)'), ('PRED_LABEL', 'Predicted'),
        ('RETROFIT_POTENTIAL', 'Actual'),
    ]
    headers = [h for _, h in display_cols]
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.style = 'Table Grid'
    table.autofit = True
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        shade_cell(cell, '2F2F2F')
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(10)
                run.font.name = 'Times New Roman'
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for i, row in enumerate(rows):
        mismatch = row['PRED_LABEL'] != row['RETROFIT_POTENTIAL']
        for j, (key, _) in enumerate(display_cols):
            cell = table.cell(i + 1, j)
            val = row[key]
            if key == 'PRED_PROBA':
                val = f"{float(val):.2f}"
            elif key == 'EFFICIENCY_GAP':
                val = f"{float(val):.0f}"
            elif key in ('PRED_LABEL', 'RETROFIT_POTENTIAL'):
                val = 'Yes' if val == '1' else 'No'
            cell.text = str(val)
            if mismatch:
                shade_cell(cell, 'F5DCDC')
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.size = Pt(10)
                    run.font.name = 'Times New Roman'
    return table


def shade_cell(cell, hex_color):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


COMPARISON_HEADER_DISPLAY = {
    'Model': 'Model',
    'CV F1-macro': 'CV F1',
    'CV ROC-AUC': 'CV AUC',
    'Test F1-macro': 'Test F1',
    'Test ROC-AUC': 'Test AUC',
    'Test PR-AUC': 'Test PR-AUC',
}
COMPARISON_COL_WIDTHS_IN = {
    # Sum must stay under ~6.2in (the floating table box width in two-column
    # mode, itself close to the full usable page width). The previous
    # widths summed to 6.6in, wider than the box: Word silently failed to
    # render the whole table rather than clipping or shrinking it, found by
    # reading the raw docx XML (the table WAS there, correctly built) after
    # it rendered as nothing in real Word.
    'Model': 1.2, 'CV F1-macro': 1.05, 'CV ROC-AUC': 1.05,
    'Test F1-macro': 0.95, 'Test ROC-AUC': 0.95, 'Test PR-AUC': 0.9,
}


def set_column_widths(table, headers, width_map, default=1.0):
    table.autofit = False
    for row in table.rows:
        for j, key in enumerate(headers):
            row.cells[j].width = Inches(width_map.get(key, default))


def add_comparison_table(doc, rows, winners):
    if rows is None:
        doc.add_paragraph("[TABLE: Model comparison. Run notebooks first.]")
        return

    best_model = winners['best_auc_model'] if winners else None
    headers = list(rows[0].keys())
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.style = 'Table Grid'

    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = COMPARISON_HEADER_DISPLAY.get(h, h)
        shade_cell(cell, '2F2F2F')
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(9)
                run.font.name = 'Times New Roman'
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, row in enumerate(rows):
        is_winner = row.get('Model') == best_model
        for j, key in enumerate(headers):
            cell = table.cell(i + 1, j)
            cell.text = str(row[key])
            if is_winner:
                shade_cell(cell, 'E8F0E3')
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.name = 'Times New Roman'
                    if is_winner:
                        run.bold = True

    set_column_widths(table, headers, COMPARISON_COL_WIDTHS_IN)
    return table


def determine_winners(comparison_rows):
    if not comparison_rows:
        return None
    ranked_auc = sorted(comparison_rows, key=lambda r: float(r['Test ROC-AUC']), reverse=True)
    ranked_f1  = sorted(comparison_rows, key=lambda r: float(r['Test F1-macro']), reverse=True)
    return {
        'ranked_auc': [{'model': r['Model'], 'test_roc_auc': float(r['Test ROC-AUC'])} for r in ranked_auc],
        'ranked_f1':  [{'model': r['Model'], 'test_f1_macro': float(r['Test F1-macro'])} for r in ranked_f1],
        'best_auc_model': ranked_auc[0]['Model'],
        'best_f1_model':  ranked_f1[0]['Model'],
        'n_models': len(comparison_rows),
    }


def set_rfonts(rFonts, name):
    """Set every ascii/hAnsi/eastAsia/cs slot to an explicit font and strip the
    theme references (asciiTheme etc). Leaving both present is what caused
    Calibri to leak through on save: python-docx's serializer favours the
    theme attribute over the explicit one once both exist on the same
    rFonts element, so the theme reference has to be removed, not just
    outranked."""
    for theme_attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
        attr_qn = qn('w:' + theme_attr)
        if rFonts.get(attr_qn) is not None:
            del rFonts.attrib[attr_qn]
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:cs'), name)


def set_style_font(style, name, size=None, color=None):
    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rpr.append(rFonts)
    set_rfonts(rFonts, name)
    if size is not None:
        style.font.size = size
    if color is not None:
        style.font.color.rgb = color


def fix_doc_defaults_font(doc, name):
    """The ultimate font fallback for any run that inherits nothing else
    (e.g. table cells under the built-in Table Grid/TableNormal styles,
    which carry no font of their own). Same theme-leak issue as styles."""
    rPrDefault = doc.styles.element.find(qn('w:docDefaults') + '/' + qn('w:rPrDefault'))
    if rPrDefault is None:
        return
    rPr = rPrDefault.find(qn('w:rPr'))
    if rPr is None:
        return
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    set_rfonts(rFonts, name)


def add_reference(doc, before, italic, after):
    p = doc.add_paragraph(style='List Paragraph')
    p.paragraph_format.first_line_indent = Pt(-18)
    r1 = p.add_run(before)
    r1.font.size = Pt(10)
    r2 = p.add_run(italic)
    r2.italic = True
    r2.font.size = Pt(10)
    r3 = p.add_run(after)
    r3.font.size = Pt(10)
    return p


def build_report(mode="full", two_column=False):
    """mode: 'full' (quality-first) or 'safe' (<=2100 words body, hard cap,
    references/title/abstract/footnotes excluded). two_column: IEEE-style
    two-column body layout, figures and tables break out to full width."""
    safe = (mode == "safe")
    suffix = ('_safe' if safe else '') + ('_2col' if two_column else '')
    OUT_PATH = f"report/25065219_report{suffix}.docx"

    doc = Document()
    set_compatibility_mode_full(doc)

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    fix_doc_defaults_font(doc, 'Times New Roman')
    set_style_font(doc.styles['Normal'], 'Times New Roman', Pt(12))
    set_style_font(doc.styles['Heading 1'], 'Times New Roman', Pt(15), RGBColor(0, 0, 0))
    set_style_font(doc.styles['Heading 2'], 'Times New Roman', Pt(13), RGBColor(0, 0, 0))
    doc.styles['Heading 1'].font.bold = True
    doc.styles['Heading 2'].font.bold = True

    comparison_rows = load_comparison_csv()
    winners = determine_winners(comparison_rows)
    bristol_rows, bristol_summary = load_bristol_case()
    fm = FootnoteManager(doc, use_endnotes=two_column)

    if bristol_summary:
        bristol_context_str = (
            f"only {(bristol_summary['bristol_accuracy'] - bristol_summary['bristol_majority_baseline_accuracy']) * 100:.1f} "
            f"points above the majority-class baseline, but with {bristol_summary['bristol_recall']:.1%} recall against "
            f"that baseline's 0%"
        )
    else:
        bristol_context_str = ""

    # ------------------------------------------------------------------
    # IEEE-style header
    # ------------------------------------------------------------------
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Predicting Retrofit Potential in UK Residential Buildings:\nA Machine Learning Classification Approach")
    run.bold = True
    run.font.size = Pt(17)

    authors = doc.add_paragraph()
    authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = authors.add_run(
        "Karan Homayounfar (25065219)\n"
        "MSc Data Science, UWE Bristol\n"
        "Machine Learning and Predictive Analytics\n"
        "Code: https://github.com/KNHNF/epc-retrofit-potential-ml"
    )
    run.font.size = Pt(11)

    doc.add_paragraph()

    # ------------------------------------------------------------------
    # Abstract
    # ------------------------------------------------------------------
    doc.add_heading("Abstract", level=2)
    winner_sentence = (
        f"{winners['best_auc_model']} achieves the highest test-set ROC-AUC of the models compared."
        if winners else
        "[RESULT: run all notebooks so the winning model can be stated correctly here.]"
    )
    abs_text = doc.add_paragraph(
        "The UK's net-zero-by-2050 target means public money for retrofit has to be aimed, not "
        "sprayed evenly across the housing stock. Most machine learning work on Energy Performance "
        "Certificate (EPC)"
    )
    if not two_column:
        fm.add(abs_text,
            "Energy Performance Certificate: a legally required rating of a UK property's energy "
            "efficiency on an A-G scale, issued whenever a home is built, sold, or rented out."
        )
    abs_text.add_run(
        " data predicts a property's current rating (A-G). I predict something more "
        "useful for that aiming problem: retrofit headroom, whether a home rated D-G has a 20-point "
        "or greater gap between its current and potential efficiency score. I compare four "
        "classifiers under nested cross-validation on 200,000 records from 2020 to 2024, then test "
        "them on held-out 2025 to 2026 data. "
        f"{winner_sentence} "
        "The positive rate also halves across that boundary, from 21.7% to 10.8%, which is "
        "why the split is by time and not at random."
        + (f" Applied to every held-out Bristol property, the same model holds "
           f"{bristol_summary['bristol_accuracy']:.1%} accuracy, {bristol_context_str}." if bristol_summary else "")
    )
    abs_text.style.font.size = Pt(11)

    doc.add_paragraph()

    index_terms = doc.add_paragraph()
    r_label = index_terms.add_run("Index Terms: ")
    r_label.bold = True
    r_label.italic = True
    r_label.font.size = Pt(11)
    r_rest = index_terms.add_run(
        "EPC, retrofit potential, binary classification, Random Forest, XGBoost, "
        "logistic regression, nested cross-validation, feature importance."
    )
    r_rest.font.size = Pt(11)

    doc.add_paragraph()

    if two_column:
        switch_columns(doc, 2)
        disable_heading_widow_control(doc)

    # ------------------------------------------------------------------
    # 1. Introduction
    # ------------------------------------------------------------------
    doc.add_heading("1. Introduction", level=1)
    p_intro1 = doc.add_paragraph()
    if safe:
        p_intro1.add_run(
            "The UK government is legally committed to net-zero greenhouse gas emissions by 2050. "
            "In the certificates I examine, 43.1% of homes are rated D or below,"
        )
        fm.add(p_intro1,
            "200,000-record training sample, 2020-2024 (Section 2)."
        )
        p_intro1.add_run(
            " and retrofitting all of them at once is neither financially nor "
            "logistically realistic [1]. So the question that matters is which homes return the "
            "most energy saved per pound of public money. Existing EPC machine learning mostly "
            "predicts the current rating label instead [2, 3]."
        )
        fm.add(p_intro1,
            "Seyedzadeh et al. [2] review current-rating prediction work directly; Pasichnyi, "
            "Wallin and Kordas [3] build city-scale archetypes rather than property-level "
            "decisions. Neither targets improvement headroom."
        )
        p_intro1.add_run(
            " That does not help much here: a home already rated C has little room to improve "
            "whatever its label. Headroom, the gap between a property's current and potential EPC "
            "score, is the thing worth flagging."
        )
    else:
        p_intro1.add_run(
            "The UK government is legally committed to net-zero greenhouse gas emissions by 2050. In the "
            "certificates I examine, 43.1% of homes are rated D or below (200,000-record training "
            "sample, 2020-2024), and "
            "retrofitting all of them at once is neither financially nor logistically realistic "
            "[1]. So the question that matters is which homes return the most energy saved per "
            "pound of public money. Existing EPC machine learning mostly answers a different question. "
            "Some of it predicts the current rating label [2]; other work builds archetypes for "
            "city-scale energy models rather than "
            "property-level decisions [3]. Neither targets improvement "
            "headroom. Predicting the rating label does not help much here, because a home already rated "
            "C has little room to improve whatever its label. Headroom is the thing worth flagging, the "
            "gap between a property's current and potential EPC score."
        )
    doc.add_paragraph("I frame this as binary classification, and use it to:")
    for aim in (
        "predict which D-G rated properties carry significant retrofit headroom (a 20-point "
        "or greater gap between current and potential efficiency)",
        "identify which characteristics actually drive that headroom, so a scheme can "
        "prioritise not just which properties, but why (Section 5.3)",
    ):
        lp = doc.add_paragraph(style='List Number')
        lp.add_run(aim)
    doc.add_paragraph(
        "Four algorithms compete: logistic regression as baseline, random forest as the main "
        "model, XGBoost as a strong benchmark, and an SVM, evaluated by nested cross-validation "
        "and a time-held-out test set, the deployment condition. A rating alone tells a "
        "homeowner where they stand, not where the money should go; headroom is the number a "
        "retrofit scheme has to rank properties by."
    )

    # ------------------------------------------------------------------
    # 2. Dataset
    # ------------------------------------------------------------------
    doc.add_heading("2. Dataset", level=1)
    doc.add_paragraph(
        "The dataset is the UK EPC Open Data published by the Ministry of Housing, Communities "
        "and Local Government [1]: domestic energy performance certificates for England and "
        "Wales, approximately 10.8 million across annual files from 2020 to 2026. After "
        "deduplication and eligibility filtering, this yields 7.25 million training-eligible "
        "certificates (2020-2024) and 2.47 million test-eligible ones (2025-2026), from which the "
        "samples below are drawn. Each certificate records construction characteristics, "
        "insulation quality, heating system, and both efficiency scores on a 1-100 scale."
    )
    p_sample = doc.add_paragraph()
    p_sample.add_run(
        "I draw a 200,000-record stratified training sample from certificates lodged in "
        "2020-2024, and a separate 50,000-record test sample from 2025-2026. Splitting by time, "
        "not at random, stops the model learning from properties assessed after the ones it "
        "predicts, exactly the situation it would face in deployment. Where a property has been "
        "assessed more than once (matched on UPRN), I keep only the most recent certificate."
    )
    fm.add(p_sample,
        "UPRN: Unique Property Reference Number, a persistent government identifier for a single "
        "property, used here to link repeat EPC assessments of the same home."
    )
    doc.add_paragraph(
        "The target is defined formally as:"
    )
    add_display_equation(
        doc, r"y = \mathbb{1}[r \in \{D,E,F,G\} \text{ and } e_{pot} - e_{cur} \geq 20]",
        two_column=two_column
    )
    doc.add_paragraph(
        "where r is current EPC rating, and e_cur and e_pot are current and potential efficiency "
        "score. That gives 21.7% positive in training and 10.8% in test (Fig. 1). The halving is "
        "worth pausing on: it is not noise, newer certificates cover homes with less headroom "
        "left, probably because more are new builds or recent refurbishments, and it is the "
        "clearest single reason a random split would have flattered the results."
    )
    add_figure(doc,
        f"{FIGURES_DIR}/01_class_balance.png",
        "Fig. 1.",
        two_column=two_column, dense=True
    )

    # ------------------------------------------------------------------
    # 2.1 Exploratory Data Analysis (C1)
    # ------------------------------------------------------------------
    doc.add_heading("2.1 Exploratory Data Analysis", level=2)
    if safe:
        p_eda1 = doc.add_paragraph()
        p_eda1.add_run(
            "Thirteen of the 41 raw fields carry missing values (Fig. 2): FLOOR_ENERGY_EFF is "
            "missing in 89.3% of records and is dropped; a second cluster is missing in 14.7% "
            "and is median/mode-imputed. A small fraction (~0.1%) carry physically impossible "
            "negative values in energy, emissions, or cost fields and are removed as sentinel or "
            "data-entry errors. A Pearson correlation check (Fig. 3) surfaces two near-collinear "
            "feature pairs,"
        )
        fm.add(p_eda1,
            "CO2 emissions per floor area correlates with total energy consumption at r=0.98, and "
            "habitable-room count with heated-room count at r=0.97 (training sample, n=200,000)."
        )
        p_eda1.add_run(
            " kept rather than dropped because split-based ensembles are not destabilised by "
            "correlated inputs the way linear coefficients are (Section 4.1). Construction age "
            "band shows a clear relationship with the target (Fig. 4): pre-1966 properties carry "
            "the highest retrofit-potential rate, consistent with Section 5.3's permutation "
            "importances."
        )
    else:
        doc.add_paragraph(
            "Thirteen of the 41 raw fields carry missing values (Fig. 2), concentrated in "
            "component-level ratings rather than spread evenly across the dataset. FLOOR_ENERGY_EFF "
            "is missing in 89.3% of records, too sparse to impute reliably, and is dropped "
            "from the feature set entirely rather than filled with a misleading default. "
            "ROOF_ENERGY_EFF is missing in 20.6% of records and is retained with an "
            "'unknown' category. A second cluster of fields, NUMBER_HABITABLE_ROOMS, "
            "MAINS_GAS_FLAG, EXTENSION_COUNT, and NUMBER_HEATED_ROOMS, is each missing in 14.7 per "
            "cent of records; these are median-imputed (numeric) or mode-imputed (categorical), "
            "since the missingness rate is low enough that imputation does not dominate the "
            "resulting distribution. TOTAL_FLOOR_AREA contains extreme outliers consistent with "
            "data entry errors (values of 0 or in the thousands of square metres for a domestic "
            "property), handled by capping at the 99th percentile before scaling. A further "
            "quality check removes the small fraction of records (~0.1%) carrying "
            "physically impossible negative values in energy consumption, CO2 emissions, or "
            "cost fields; these are sentinel or data-entry errors rather than real measurements, "
            "so the affected records are dropped, consistent with the efficiency range filter."
        )
        doc.add_paragraph(
            "A Pearson correlation matrix over the numeric features (Fig. 3) surfaces two "
            "near-collinear pairs: CO2 emissions per floor area against total energy consumption "
            "(r=0.98) and habitable-room count against heated-room count (r=0.97), on the 200,000-row "
            "training sample. Neither pair is dropped. Split-based ensembles choose one correlated "
            "feature over another without the estimate instability that collinearity causes for a "
            "linear model's coefficients, which is one concrete reason Logistic Regression is kept "
            "as a baseline rather than promoted to the main model (Section 4.1)."
        )
        doc.add_paragraph(
            "Construction age band shows a clear relationship with the target (Fig. 4): pre-1966 "
            "properties, built before modern insulation standards became common, carry the highest "
            "retrofit-potential rate of any age band. This is a useful sanity check on the target "
            "definition itself, not just a modelling input: it confirms the 20-point efficiency-gap "
            "threshold is picking out properties that match independent domain expectations about "
            "where retrofit headroom actually sits, and it is consistent with construction age band "
            "ranking highly in the Random Forest permutation importances reported in Section 5.3."
        )
    # Floating figures cross both columns without a section break, so
    # these three no longer need any column-switch wrapping at all.
    add_figure(doc,
        f"{FIGURES_DIR}/02_missing_values.png",
        "Fig. 2.",
        width=5.5, two_column=two_column, dense=True
    )
    add_figure(doc,
        f"{FIGURES_DIR}/04_correlation_heatmap.png",
        "Fig. 3.",
        width=5.5, two_column=two_column, dense=True
    )
    add_figure(doc,
        f"{FIGURES_DIR}/06_retrofit_rate_by_age.png",
        "Fig. 4.",
        width=5.5, two_column=two_column, dense=False
    )

    # ------------------------------------------------------------------
    # 3. Problem Definition
    # ------------------------------------------------------------------
    doc.add_heading("3. Problem Definition", level=1)
    doc.add_paragraph(
        "This is a multivariate supervised binary classification problem. Features are numerical "
        "(current energy efficiency, floor area, CO2 emissions per floor area, heating and "
        "hot-water costs, room counts), ordinal (rating encoded A=6 to G=0; component ratings "
        "Very Good=5 to N/A=0), and nominal (property type, built form, wall type, tenure, mains "
        "gas flag). All are observable at assessment time, and none are derived from the target."
    )
    doc.add_paragraph(
        "Potential energy efficiency and potential rating are excluded from the feature set. "
        "Including either would be target leakage: the efficiency gap is computed directly from "
        "potential efficiency, so a model with access to it would trivially predict rather than "
        "learn from physical characteristics. Current energy efficiency is retained since it is "
        "fully observable and does not directly encode the potential score."
    )

    # ------------------------------------------------------------------
    # 4. Algorithm Selection and Methodology
    # ------------------------------------------------------------------
    doc.add_heading("4. Algorithm Selection and Methodology", level=1)
    p_impl = doc.add_paragraph()
    if safe:
        p_impl.add_run(
            "Models are implemented in Python via scikit-learn [4]. Logistic Regression, Random "
            "Forest, and k-fold cross-validation were taught here; XGBoost and nested "
            "cross-validation were not, and both are justified below rather than used unremarked."
        )
        fm.add(p_impl,
            "I chose to go beyond the taught syllabus, not to avoid it: every taught method still "
            "appears (Logistic Regression as the interpretable baseline, Random Forest as the main "
            "model, standard k-fold CV as the concept nested CV extends), and each self-directed "
            "addition is justified against a specific limitation of the taught alternative, not "
            "used because it sounded more advanced."
        )
    else:
        p_impl.add_run(
            "I implement everything in Python with scikit-learn [4], plus XGBoost's own library "
            "for the gradient-boosting model in 4.3 [5]."
        )
    doc.add_heading("4.1 Logistic Regression", level=2)
    p_lr = doc.add_paragraph()
    if safe:
        p_lr.add_run(
            "Logistic Regression serves as the baseline. It is interpretable via its "
            "coefficients, provides well-calibrated probabilities, and sets a minimum "
            "performance bar the other three models must exceed."
        )
        fm.add(p_lr,
            "Ng and Jordan [6] show that discriminative classifiers reach their asymptotic error "
            "with fewer training examples than generative alternatives such as Naive Bayes, "
            "which is why LR rather than NB is the baseline here."
        )
        p_lr.add_run(" The near-collinear pairs found in Section 2.1 are a further reason.")
        fm.add(p_lr,
            "They would inflate LR's coefficient variance without necessarily hurting its "
            "predictions, a problem split-based ensembles do not share."
        )
    else:
        p_lr.add_run(
            "Logistic Regression serves as the baseline. Ng and Jordan [6] showed that "
            "discriminative classifiers reach their asymptotic error with fewer training examples "
            "than generative models (e.g. Naive Bayes), justifying LR over NB on a dataset of "
            "this size. LR is interpretable via its coefficients, provides well-calibrated "
            "probabilities, and sets a minimum performance bar that more complex models must exceed. "
            "The near-collinear feature pairs identified in Section 2.1 are a further reason it "
            "stays a baseline rather than the main model: they would inflate LR's coefficient "
            "variance without necessarily hurting its predictions, a problem split-based ensembles "
            "do not share."
        )
    doc.add_heading("4.2 Random Forest", level=2)
    p_rf = doc.add_paragraph()
    if safe:
        p_rf.add_run(
            "Random Forest is my main model. Bagging many decorrelated trees over random feature "
            "subsets cuts variance without adding bias [7], which is exactly what a single "
            "Decision Tree lacks, so I reject a standalone Decision Tree outright."
        )
        fm.add(p_rf,
            "Unpruned trees overfit badly, and even pruned ones still lose to the ensemble on "
            "both bias-variance tradeoff and generalisation [8]."
        )
        p_rf.add_run(
            " For feature importance I use permutation importance rather than the default "
            "impurity-based score (Mean Decrease in Impurity, MDI)."
        )
        fm.add(p_rf,
            "MDI is known to be unreliable across features of mixed cardinality [9], and this "
            "feature set mixes one-hot categoricals with continuous numerics."
        )
    else:
        p_rf.add_run(
            "Random Forest is my main model. Bagging many decorrelated trees over random feature "
            "subsets cuts variance without adding bias [7], which is exactly what a single decision "
            "tree lacks. I reject a standalone Decision Tree outright: unpruned trees overfit badly, "
            "and even pruned ones still lose to the ensemble on both bias-variance tradeoff and "
            "generalisation [8]. For feature importance I use permutation importance rather than the "
            "default impurity-based score (Mean Decrease in Impurity, MDI), because MDI is known to be "
            "unreliable across features of mixed cardinality [9], and my feature set mixes one-hot "
            "categoricals with continuous numerics."
        )
    doc.add_heading("4.3 XGBoost", level=2)
    p_xgb = doc.add_paragraph()
    if safe:
        p_xgb.add_run(
            "XGBoost was not covered on this module. I add it because boosting is structurally "
            "different from bagging, testing whether 4.2's result is a property of tree "
            "ensembles generally, or specific to bagging."
        )
        fm.add(p_xgb,
            "Bagging (Random Forest) and boosting (XGBoost) sit on opposite ends of the "
            "bias-variance tradeoff taught in 4.2: bagging averages many independent, "
            "low-bias/high-variance trees to cut variance, boosting builds trees sequentially, "
            "each correcting the previous ensemble's residual error, trading some variance "
            "reduction for lower bias. I read Chen and Guestrin [5] independently to justify using "
            "it here, since the module itself did not cover gradient boosting. It extends gradient "
            "boosting with second-order Taylor approximations of the loss, column subsampling, and "
            "L1/L2 regularisation, and is one of the strongest published performers on tabular "
            "benchmarks [5]."
        )
        p_xgb.add_run(
            " I expected it to beat Random Forest outright; Section 6 covers why it didn't. "
            "Class imbalance is handled via scale_pos_weight, which weights the positive class "
            "inversely to its frequency."
        )
    else:
        p_xgb.add_run(
            "XGBoost is my comparison against Random Forest for tree-based ensembles. It extends "
            "gradient boosting with second-order Taylor approximations of the loss, column "
            "subsampling, and L1/L2 regularisation, and is consistently one of the strongest "
            "published performers on tabular benchmarks [5]. I expected it to beat Random Forest outright; Section 6 covers why it "
            "didn't. Here, scale_pos_weight handles the class imbalance by weighting the positive "
            "class inversely to its frequency."
        )
    doc.add_heading("4.4 Reference Models: SVM and kNN", level=2)
    p_svm = doc.add_paragraph()
    if safe:
        p_svm.add_run(
            "SVM and kNN sit here as reference points, not primary candidates: Table 1 reports "
            "SVM's full metrics, and kNN was never fitted."
        )
        fm.add(p_svm,
            "LinearSVC has no native probability output, hence CalibratedClassifierCV (Platt "
            "scaling) to get scores for ROC-AUC [10]. Past roughly 10-15 dimensions, the distance "
            "to the nearest neighbour converges toward the distance to the farthest one [11]; at "
            "51 dimensions here, kNN would not have been a fair comparison, not a broken one. "
            "Beyer et al.'s result [11] is derived for continuous i.i.d. features, and this "
            "51-dimensional space is mostly one-hot categoricals plus a handful of continuous "
            "fields, not an exact match, so this is a heuristic argument, not a strict bound."
        )
    else:
        p_svm.add_run(
            "SVM (linear kernel, Platt-calibrated via CalibratedClassifierCV since LinearSVC has "
            "no native probability output [10]) and kNN sit in this report as reference points, "
            "not primary candidates alongside Random Forest and XGBoost. Table 1 reports SVM's "
            "full test-set metrics for completeness. kNN was never fitted: past roughly 10-15 "
            "dimensions the distance to the nearest neighbour converges toward the distance to "
            "the farthest one, and Euclidean distance stops carrying useful discriminative signal "
            "[11]. At 51 dimensions here, kNN would not have been a fair comparison, not a broken "
            "one, though I should flag that Beyer et al.'s result [11] is derived for continuous "
            "i.i.d. features, and this feature space is mostly one-hot categoricals plus a handful "
            "of continuous fields, not an exact match for that assumption. It is a heuristic "
            "argument for excluding kNN, not a strict mathematical guarantee that it would fail here."
        )
    doc.add_heading("4.5 Class Imbalance and Evaluation Metrics", level=2)
    p_metrics = doc.add_paragraph()
    p_metrics.add_run(
        "Every model uses class_weight='balanced' (or its equivalent) to correct for the 78:22 "
        "imbalance in training data. I report F1-macro and ROC-AUC throughout."
    )
    fm.add(p_metrics,
        "ROC-AUC: area under the receiver operating characteristic curve, a threshold-independent "
        "ranking score (0.5 is no better than chance, 1.0 is perfect). PR-AUC: the same idea for "
        "the precision-recall curve, more informative than ROC-AUC on an imbalanced target since "
        "its baseline is the positive rate, not 0.5. F1-macro: the unweighted average of the "
        "F1 score (precision/recall balance) across both classes, so the minority class counts as "
        "much as the majority one."
    )
    p_metrics.add_run(
        " Accuracy is not a primary metric here: a model that just predicts the majority class "
        "every time would score 78.3% on training data without learning anything."
    )
    doc.add_heading("4.6 Nested Cross-Validation and Temporal Test Split", level=2)
    p_ncv = doc.add_paragraph()
    p_ncv.add_run(
        "Standard k-fold cross-validation was taught here; nested cross-validation extends it "
        "with an inner loop, used because a single k-fold split has a specific flaw once it also "
        "chooses hyperparameters."
    )
    fm.add(p_ncv,
        "If the same k-fold split both selects hyperparameters (by picking whichever setting "
        "scores best on the held-out folds) and then reports that score as the model's expected "
        "performance, the reported score is optimistically biased: it has effectively been chosen "
        "for scoring well on those exact folds [12]. Nested CV separates the two jobs into an "
        "outer loop (scores generalisation) and an inner loop (selects hyperparameters), so the "
        "score reported was never used to pick the setting being scored."
    )
    p_ncv.add_run(
        " A 5-fold stratified outer loop estimates generalisation performance; a 3-fold stratified "
        "inner loop selects hyperparameters. The final test evaluation then uses the temporally "
        "held-out 2025-2026 sample, never seen during training or tuning, a stricter check than "
        "nested CV alone provides."
    )

    # ------------------------------------------------------------------
    # 5. Results
    # ------------------------------------------------------------------
    doc.add_heading("5. Results", level=1)
    doc.add_heading("5.1 Model Comparison", level=2)
    doc.add_paragraph(
        "Table 1 presents the nested cross-validation scores on the training set and the final "
        "held-out test set scores for all four models."
    )

    cap1 = doc.add_paragraph("Table 1. Model comparison: nested CV and test set performance.")
    cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap1.runs[0].italic = True
    cap1.runs[0].font.size = Pt(10)

    table1_img = f"{FIGURES_DIR}/table1_model_comparison.png"
    if two_column and os.path.exists(table1_img):
        # Real Word tables cannot float across both columns in Word (three
        # separate mechanisms tried and confirmed broken in real Word:
        # section break, DrawingML text box, native w:tblpPr), so the
        # two-column variant renders the same real data as an image
        # instead, using the picture float that already works. The
        # single-column variant keeps a real, editable Word table below.
        # Caption already added above (matches Table 2's pattern), so this
        # picture float gets a single space, not a duplicate caption.
        add_floating_picture(doc, table1_img, " ")
    else:
        add_comparison_table(doc, comparison_rows, winners)
        doc.add_paragraph()

    if winners:
        auc_order = ", ".join(
            f"{m['model']} ({m['test_roc_auc']:.4f})" for m in winners['ranked_auc']
        )
        if safe:
            result_sentence = (
                f"On the held-out test set, models rank by ROC-AUC as follows: {auc_order}. "
                f"{winners['best_auc_model']} scores highest, despite not leading on "
                "cross-validation (Section 6)."
            )
        else:
            result_sentence = (
                f"On the held-out test set the models rank by ROC-AUC as: {auc_order}. "
                f"{winners['best_auc_model']} comes top. The part I did not expect is that this is "
                "not the model that scored best in cross-validation. CV runs on the 2020-2024 "
                "training years; the test set is a later, structurally different period, and the two "
                "do not agree (see temporal shift below)."
            )
    else:
        result_sentence = (
            "[RESULT: run all notebooks and re-generate this report so the ranked "
            "comparison can be stated here from the actual saved metrics.]"
        )
    p_res = doc.add_paragraph()
    if safe:
        p_res.add_run(f"Every model clears the no-skill baseline comfortably. {result_sentence.rstrip()}")
        fm.add(p_res,
            "No-skill baseline: ROC-AUC 0.5, PR-AUC equal to the positive prevalence."
        )
        p_res.add_run(
            " The CV-to-test drop shows up for all four and traces back to the temporal shift "
            "from Section 2: test years hold fewer high-headroom homes, so the task is harder "
            "there, matching deployment rather than a fault in the models."
        )
    else:
        p_res.add_run(
            "Every model clears the no-skill baseline comfortably (ROC-AUC 0.5, PR-AUC equal to the "
            "positive prevalence). "
            f"{result_sentence} "
            "The CV-to-test drop shows up for all four and traces straight back to the temporal shift "
            "from Section 2: the test years hold far fewer high-headroom homes, so the task is harder "
            "there. That is what deployment on future data looks like, so I read the drop as realistic, "
            "not as a fault in the models."
        )

    p_pr = doc.add_paragraph()
    if safe:
        p_pr.add_run(
            "In raw terms, Random Forest scores 92.0% accuracy, 58.7% precision, 86.4% recall. "
            "Accuracy overstates this: guessing the majority class scores 89.2% free at this "
            "test set's 10.8% positive rate,"
        )
        fm.add(p_pr,
            "Real numbers from rf_test_preds.npy against y_test.npy, not estimated."
        )
        p_pr.add_run(
            " so recall does the real work, and F1-macro (Table 1) balances both rather than "
            "reporting either alone. All four curves sit close together (Fig. 5): the PR curves "
            "only really separate as recall approaches 1.0, where Random Forest holds precision "
            "longest."
        )
    else:
        p_pr.add_run(
            "In raw terms, Random Forest scores 92.0% test accuracy, 58.7% precision, and 86.4% "
            "recall on the held-out set. Accuracy alone overstates the result here: a model that "
            "always guessed the majority class would score 89.2% for free at this test set's "
            "10.8% positive rate, so recall is the number actually doing the work, not accuracy. "
            "F1-macro (Table 1) is reported as the headline metric instead of precision or recall "
            "alone because it balances both across both classes, and does not privilege whichever "
            "one a specific threshold happens to favour. Visually, all four models' curves sit "
            "close together (Fig. 5), the ROC curves stay near the top-left corner across most "
            "thresholds, and the PR curves only really pull apart as recall approaches 1.0, where "
            "Random Forest holds its precision longest."
        )

    add_figure(doc,
        f"{FIGURES_DIR}/all_models_roc_pr.png",
        "Fig. 5.",
        two_column=two_column, dense=True
    )

    doc.add_heading("5.2 Statistical Significance: McNemar's Test", level=2)
    if safe:
        p_mcn = doc.add_paragraph()
        p_mcn.add_run(
            "McNemar's test was applied to four model pairs to assess whether error "
            "distributions are statistically distinct [13], the right choice for a single "
            "train/test split."
        )
        fm.add(p_mcn,
            "It has acceptably low Type I error here, as opposed to designs with repeated "
            "resampling, where a 5x2cv test is recommended instead [14]; the Diebold-Mariano test "
            "is for regression and does not apply. Two alternatives were considered and rejected: "
            "a paired t-test on the nested CV fold scores would violate the independence "
            "assumption, since outer folds share overlapping training data through the inner loop; "
            "a bootstrap confidence interval on the accuracy difference is a reasonable alternative "
            "but does not test the specific hypothesis McNemar does, that the two classifiers' "
            "errors are exchangeable, so both are non-fits, not just McNemar being picked by default."
        )
        p_mcn.add_run(
            " Every pair differs significantly (p < 0.0001), including SVM vs Random Forest, "
            "the pair with the smallest McNemar test statistic"
        )
        fm.add(p_mcn,
            "Smallest test statistic, not the closest pair on ROC-AUC: that is Random Forest vs "
            "XGBoost (Table 1), a different comparison. McNemar's statistic measures how "
            "one-sided the disagreements are (each model right where the other is wrong), not "
            "how large the accuracy gap is."
        )
        p_mcn.add_run(
            ". With 50,000 test rows this has power to flag small differences: the models make "
            "different mistakes on different properties, not that any single decision would "
            "change. Table 1's ROC-AUC and F1-macro gaps guide practical importance better."
        )
    else:
        doc.add_paragraph(
            "McNemar's test was applied to four model pairs (Random Forest vs Logistic Regression, "
            "XGBoost vs Logistic Regression, XGBoost vs Random Forest, and SVM vs Random Forest) to "
            "assess whether classification error distributions are statistically distinct "
            "[13]. It is the right choice here: of five candidate significance tests for "
            "comparing classifiers, it has acceptably low Type I error specifically for the single "
            "train/test split design used in this report, as opposed to designs involving repeated "
            "resampling, where a 5x2cv test is recommended instead [14]. The test is "
            "appropriate for binary classifiers evaluated on the same test "
            "instances; the Diebold-Mariano test is for regression and is not applicable here. Every "
            "pair differs significantly (p < 0.0001), including SVM vs Random Forest, the pair with "
            "the smallest McNemar test statistic, not the same as the closest pair on ROC-AUC "
            "(that is Random Forest vs XGBoost, Table 1): McNemar's statistic measures how one-sided "
            "the disagreements are, not the size of the accuracy gap. "
            "With a 50,000-row test set, "
            "McNemar's test has enough statistical power to detect small differences in error patterns "
            "as significant; this establishes that the models make different mistakes on different "
            "properties, not that the difference is large enough to matter for any single "
            "property-level decision. The ROC-AUC and F1-macro gaps in Table 1 are a better guide to "
            "practical importance than the p-values alone."
        )

    add_figure(doc,
        f"{FIGURES_DIR}/rf_permutation_importances.png",
        "Fig. 6.",
        two_column=two_column, dense=True
    )

    doc.add_heading("5.3 Feature Importance", level=2)
    p_fi = doc.add_paragraph()
    if safe:
        p_fi.add_run(
            "Random Forest permutation importances (Fig. 6) rank current energy efficiency score "
            "as the strongest predictor by a wide margin, followed by CO2 emissions per floor "
            "area and heating cost. Construction age band also ranks highly: pre-1966 stock "
            "likely carries more headroom (Fig. 4)."
        )
        fm.add(p_fi,
            "This report does not include a separate age-stratified breakdown to verify that "
            "mechanism directly, so this is a proposed explanation, not a demonstrated one."
        )
        p_fi.add_run(
            " Wall type (cavity vs solid) also contributes: EPC's methodology rates uninsulated "
            "solid-wall construction as poor by design, so solid-wall properties systematically "
            "carry more headroom, a property of how EPC scores are assessed, not an artefact of "
            "the model."
        )
    else:
        p_fi.add_run(
            "Feature importance analysis from Random Forest permutation importances (Fig. 6) "
            "shows that current energy efficiency score is the single strongest predictor, followed by CO2 "
            "emissions per floor area and heating cost. Construction age band also ranks highly: "
            "older properties, particularly pre-1966 stock built before modern insulation standards, "
            "typically carry more retrofit headroom (Fig. 4); this report does not "
            "include a separate age-stratified breakdown to verify that mechanism directly, so this "
            "is a proposed explanation, not a demonstrated one. "
            "Wall type (cavity vs solid) contributes meaningfully: EPC's assessment methodology rates "
            "uninsulated solid-wall construction as poor by design, so solid-wall properties "
            "systematically score lower on baseline efficiency and carry more improvement headroom "
            "than cavity-wall equivalents. This is a property of how EPC scores are assessed, not an "
            "artefact of the model."
        )

    doc.add_heading("5.4 Ablation: Naive Structural Baseline", level=2)
    naive = load_naive_baseline()
    best_row = None
    if winners and comparison_rows:
        best_row = next((r for r in comparison_rows if r['Model'] == winners['best_auc_model']), None)

    if naive and best_row:
        best_auc = float(best_row['Test ROC-AUC'])
        best_f1  = float(best_row['Test F1-macro'])
        best_pr  = float(best_row['Test PR-AUC'])
        auc_gap = best_auc - naive['test_roc_auc']
        f1_gap  = best_f1  - naive['test_f1_macro']
        pr_gap  = best_pr  - naive['test_pr_auc']
        if safe:
            p_naive = doc.add_paragraph()
            p_naive.add_run(
                "A naive Logistic Regression restricted to CURRENT_ENERGY_EFFICIENCY and "
                "CURRENT_ENERGY_RATING, fitted under the same protocol, tests how much of the "
                "result these two fields explain alone."
            )
            fm.add(p_naive,
                "CURRENT_ENERGY_EFFICIENCY is retained as a feature even though a low current "
                "score structurally leaves more numerical room for a large gap, since potential "
                "efficiency is capped near 100."
            )
            p_naive.add_run(
                f" It reaches a test ROC-AUC of {naive['test_roc_auc']:.4f}, only {auc_gap:.4f} "
                f"below {winners['best_auc_model']}'s {best_auc:.4f}. On the class-imbalance-sensitive "
                f"metrics the gap is larger: {f1_gap:.4f} F1-macro and {pr_gap:.4f} PR-AUC, showing "
                "the structural and construction-age features contribute real value beyond the "
                "current-to-potential arithmetic relationship."
            )
        else:
            doc.add_paragraph(
                "RETROFIT_POTENTIAL is derived from the gap between potential and current EPC "
                "efficiency scores. POTENTIAL_ENERGY_EFFICIENCY is excluded from the feature set "
                "(Section 3), but CURRENT_ENERGY_EFFICIENCY is included, and a low current score "
                "structurally leaves more numerical room for a large gap, since potential efficiency "
                "is capped near 100. To test how much of the result above this relationship explains "
                "on its own, a Logistic Regression model restricted to only CURRENT_ENERGY_EFFICIENCY "
                "and CURRENT_ENERGY_RATING was fitted under the same nested cross-validation protocol "
                f"as the main baseline. This naive model reaches a test ROC-AUC of "
                f"{naive['test_roc_auc']:.4f}, only {auc_gap:.4f} below {winners['best_auc_model']}'s "
                f"{best_auc:.4f}, which explains why current efficiency dominates the permutation "
                "importance ranking above. On F1-macro and PR-AUC, the metrics more informative under "
                f"class imbalance, the gap is substantially larger: {naive['test_f1_macro']:.4f} versus "
                f"{best_f1:.4f} test F1-macro ({f1_gap:.4f} difference) and {naive['test_pr_auc']:.4f} "
                f"versus {best_pr:.4f} test PR-AUC ({pr_gap:.4f} difference). The structural, tenure, "
                "and construction-age features therefore contribute real discriminative value beyond "
                "the current-to-potential arithmetic relationship, even though ROC-AUC alone "
                "understates that contribution."
            )
    else:
        doc.add_paragraph(
            "[ABLATION: run the naive baseline cell in 05_Comparison_Evaluation.ipynb "
            "and re-generate this report so this section can be stated from the actual saved "
            "metrics.]"
        )

    doc.add_heading("5.5 Calibration", level=2)
    doc.add_paragraph(
        "Calibration checks whether a model's reported probability matches the true observed "
        "frequency of the positive class, distinct from ranking ability (ROC-AUC), and matters "
        "for policy use, where a threshold needs to be trustworthy, not just the ranking. All "
        "four curves sit below the diagonal (Fig. 7): every model over-predicts risk, a property "
        "flagged at 70% is genuinely high-potential closer to 40-50% of the time. Logistic "
        "Regression is most over-confident; Random Forest and SVM track closest to each other "
        "and are the least mis-calibrated."
    )
    add_figure(doc,
        f"{FIGURES_DIR}/calibration_curves.png",
        "Fig. 7.",
        two_column=two_column, dense=False
    )

    # ------------------------------------------------------------------
    # 6. Discussion
    # ------------------------------------------------------------------
    doc.add_heading("6. Discussion", level=1)
    if winners:
        lead_model = winners['best_auc_model']
        if safe:
            lead_sentence = (
                f"{lead_model} achieves the strongest test-set performance of the four models "
                "compared."
            )
        else:
            lead_sentence = (
                f"{lead_model} comes out ahead on every test-set metric of the four models "
                "compared."
            )
    else:
        lead_sentence = "[RESULT: state which model leads once all notebooks have been run.]"
    if safe:
        doc.add_paragraph(
            "Machine learning can identify high-retrofit-potential properties from observable "
            f"EPC characteristics with meaningful accuracy. {lead_sentence} For policy use, the "
            "choice between the top two tree-based models depends on interpretability: Random "
            "Forest's permutation importances are straightforward to communicate to "
            "policymakers, while XGBoost's gain-based importances can be biased toward "
            "high-cardinality features."
        )
    else:
        doc.add_paragraph(
            "Machine learning finds high-retrofit-potential properties in EPC data with real "
            "accuracy, not just a marginal lift over guessing. "
            f"{lead_sentence} "
            "For policy use, picking between the two tree-based models comes down to how you need "
            "to explain the result: Random Forest's permutation importances are easy to hand to a "
            "policymaker, while XGBoost's gain-based importances skew toward high-cardinality "
            "features and are harder to defend in plain language."
        )
    if safe:
        doc.add_paragraph(
            "This connects back to why Random Forest was the main model, not just the winner: "
            "bagging trades training-set fit for lower variance [7], which held up on data "
            "XGBoost's flexibility had never seen. I read the gap as consistent with that "
            "mechanism, not proven, since nothing here isolates variance from the algorithms' "
            "other differences."
        )
        doc.add_paragraph(
            "The temporal shift (21.7% to 10.8% positive) matters for deployment: future stock "
            "will have proportionally fewer high-retrofit candidates. Relative ranking still "
            "holds, but a fixed threshold will lose recall over time, so annual re-calibration "
            "against Fig. 7 should be recall-oriented, not accuracy-oriented."
        )
    else:
        doc.add_paragraph(
            "This connects back to why Random Forest was the main model in the first place, not "
            "just the eventual winner. Bagging trades a little training-set fit for lower "
            "variance [7], which is exactly what would let it hold up better on a test period "
            "XGBoost's extra flexibility had never seen. I read the CV-to-test gap as consistent "
            "with that mechanism, not proof of it: nothing in this report isolates variance from "
            "the other ways the two algorithms differ, so it stays an untested explanation, "
            "argued from the same bias-variance logic that motivated the model choice in Section 4.2."
        )
        doc.add_paragraph(
            "The positive rate drops from 21.7% to 10.8% across the temporal split, and that "
            "matters for deployment. A model trained on 2020-2024 data and pointed at future "
            "assessments will meet a stock with proportionally fewer high-retrofit candidates. "
            "That doesn't break the model. What a prioritisation tool needs is the relative "
            "ranking of properties, and that still holds. What it does mean is that a fixed "
            "probability threshold loses recall over time, quietly, without the model itself "
            "changing. I would re-calibrate the threshold annually against the reliability "
            "diagrams in Fig. 7, biased toward recall rather than accuracy, rather than trust a "
            "threshold picked once at launch."
        )
    doc.add_paragraph("There are five main limitations.")
    if safe:
        limitation_items = [
            "Data quality. The true EPC error rate is estimated at 36-62% once assessor "
            "disagreement is accounted for [15], which affects WALL_TYPE, engineered from the "
            "same unreliable free-text fields.",
            "Target simplification. The binary target collapses heterogeneous properties: a "
            "20-point gap means different things in a rural solid-wall home versus an urban flat.",
            "Sample size. This uses 200,000 of 7.25 million eligible records; full-data training "
            "may improve minority-class recall.",
            "Spatial features. Regional and deprivation variables enter as nominal categories, not "
            "spatial models.",
            "Single temporal split, not a rolling backtest. Evaluation uses one static "
            "2020-2024/2025-2026 split, not walk-forward retraining across several periods, so "
            "it cannot show whether performance is stable or drifting over time.",
        ]
    else:
        limitation_items = [
            "Data quality. The EPC database has documented quality issues: 27% of "
            "open-data EPCs carry at least one flag suggesting an error, and the true error rate "
            "is estimated at 36% to 62% once assessor disagreement on parameters such as "
            "wall type and built form is accounted for [15]. This directly "
            "affects the WALL_TYPE feature engineered in this pipeline, which is derived from the "
            "same free-text description fields identified there as unreliable.",
            "Target simplification. The binary target collapses heterogeneous properties: a "
            "20-point gap in a rural solid-wall property has different policy implications from "
            "the same gap in an urban flat.",
            "Sample size. This analysis uses a sample of 200,000 records rather than the full "
            "7.25 million training records; while sufficient for credible results, full-data "
            "training may improve recall on the minority class.",
            "Spatial features. Regional clustering and local authority deprivation indices are "
            "used only as nominal categories and may benefit from spatial modelling approaches.",
            "Single temporal split, not a rolling backtest. The final evaluation here uses one "
            "static split, train on 2020-2024, test once on 2025-2026, rather than a repeated "
            "walk-forward (expanding-window) evaluation across several successive retraining "
            "points, which is closer to how a model actually gets re-evaluated in deployment. "
            "Nested cross-validation avoids the optimistic bias of tuning and scoring on the same "
            "split, but that is a different problem from this one: a single train/test split still "
            "cannot show whether performance is stable, improving, or drifting across successive "
            "periods, only that it holds on this one boundary.",
        ]
    for item in limitation_items:
        lp = doc.add_paragraph(style='List Number')
        lp.add_run(item)

    # ------------------------------------------------------------------
    # 7. Real-World Application: A Bristol Case Study
    # ------------------------------------------------------------------
    doc.add_heading("7. Real-World Application: A Bristol Case Study", level=1)
    if bristol_rows and bristol_summary:
        n = bristol_summary['n_bristol_test_properties']
        acc = bristol_summary['bristol_accuracy']
        maj_base = bristol_summary['bristol_majority_baseline_accuracy']
        recall = bristol_summary['bristol_recall']
        precision = bristol_summary['bristol_precision']
        f1 = bristol_summary['bristol_f1']
        pos_rate = bristol_summary['bristol_positive_rate']
        rest_rate = bristol_summary['rest_of_test_positive_rate']
        gap_pts = bristol_summary['positive_rate_gap_pts']
        p_value = bristol_summary['positive_rate_pvalue']
        p_bristol = doc.add_paragraph()
        if safe:
            p_bristol.add_run(
                f"To check this is more than a benchmark exercise, I ran the trained Random "
                f"Forest on every Bristol, City of certificate in the held-out 2025-2026 test "
                f"set: {n:,} properties never seen during training or tuning."
            )
            fm.add(p_bristol,
                "Matched by UPRN against the open EPC register's own local-authority field, not "
                "a synthetic or hand-picked set."
            )
            p_bristol.add_run(
                f" Accuracy is {acc:.1%}, not meaningful alone: guessing \"not high potential\" "
                f"always scores {maj_base:.1%} free, since only {pos_rate:.1%} of Bristol "
                f"properties are genuinely high-potential. Recall shows what accuracy hides: "
                f"{recall:.1%} of true positives flagged (precision {precision:.1%}, F1 "
                f"{f1:.2f}) against 0% for that baseline. Bristol's {pos_rate:.1%} positive rate "
                f"sits {abs(gap_pts):.1f} points below the rest ({rest_rate:.1%}), significant "
                f"(z-test, p={p_value:.3f})"
            )
            fm.add(p_bristol,
                "Likely a real feature of Bristol's housing stock rather than a model artefact, "
                "given the model was never trained or tuned on Bristol-specific data."
            )
            p_bristol.add_run(".")
        else:
            p_bristol.add_run(
                f"A benchmark result is not the same as a working tool, so I ran the trained "
                f"Random Forest on every Bristol, City of certificate in the held-out 2025-2026 "
                f"test set: {n:,} real properties the model never saw during training or "
                f"hyperparameter tuning, identified by matching UPRN against the open EPC "
                f"register's own local-authority field, not a synthetic or hand-picked "
                f"demonstration set. Accuracy on this single-city slice is {acc:.1%}, close to "
                f"the overall test-set figure in Table 1, but accuracy on its own overstates how "
                f"good this is: only {pos_rate:.1%} of Bristol properties are genuinely "
                f"high-potential, so a model that always predicted \"not high potential\" would "
                f"score {maj_base:.1%} by construction, without ever finding a single property "
                f"worth retrofitting. Recall is the metric that actually distinguishes the two: "
                f"the trained model correctly flags {recall:.1%} of Bristol's true high-potential "
                f"properties (precision {precision:.1%}, F1 {f1:.2f}) against 0% recall for that "
                f"trivial baseline. Bristol's {pos_rate:.1%} positive rate "
                f"sits {abs(gap_pts):.1f} points below the {rest_rate:.1%} rate in the rest of "
                f"the test set, and at this sample size that gap is statistically significant "
                f"(two-proportion z-test, p={p_value:.3f}), not noise. I read that as a real "
                f"property of Bristol's housing stock rather than a model artefact, since the "
                f"model was never trained or tuned on anything Bristol-specific, but the "
                f"headline result here is not that the rates match exactly, it is that recall "
                f"holds up on a city the model has never seen."
            )
        district_summary_path = f"{DATA_DIR}/bristol_district_summary.csv"
        if os.path.exists(district_summary_path) and os.path.exists(
            f"{FIGURES_DIR}/07_bristol_district_map.png"
        ):
            p_map = doc.add_paragraph()
            p_map.add_run(
                "Predicted rate varies by district (Fig. 8), from BS1 (3.5%, lowest) to BS15 "
                "(21.9%, highest); marker position is an approximate centroid, not a real "
                "boundary."
            )
            add_figure(doc,
                f"{FIGURES_DIR}/07_bristol_district_map.png",
                "Fig. 8.",
                two_column=two_column, dense=True
            )
        cap_bristol = doc.add_paragraph()
        cap_bristol.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap_bristol.add_run(
            "Table 2. Eight Bristol test-set properties: model output vs. actual label."
        )
        cap_run.italic = True
        cap_run.font.size = Pt(10)
        cap_run.font.name = 'Times New Roman'
        fm.add(cap_bristol,
            f"Stratified sample from the {n:,} Bristol test-set properties: 3 correctly "
            "predicted positives, 3 correctly predicted negatives, and 2 genuine model errors, "
            "each group sampled at random with a fixed seed (42), not hand-picked."
        )
        table2_img = f"{FIGURES_DIR}/table2_bristol.png"
        if two_column and os.path.exists(table2_img):
            # Same real-table-cannot-float-in-Word limitation as Table 1;
            # caption already added above (with its footnote), so this
            # picture float is captioned separately with a single space to
            # avoid an empty run, not duplicating the real caption text.
            add_floating_picture(doc, table2_img, " ")
        else:
            add_bristol_table(doc, bristol_rows)
        doc.add_paragraph()
        p_use = doc.add_paragraph()
        if safe:
            p_use.add_run(
                "This is what the tool is for: a buyer, landlord, or council screens a shortlist "
                "by postcode for a ranked headroom probability. Not a valuation tool: the two "
                "shaded rows are genuine model errors, and neither purchase price nor retrofit "
                "cost is estimated here,"
            )
            fm.add(p_use,
                "A real investment decision still needs a separate estimate of retrofit cost "
                "against likely value uplift; Land Registry Price Paid Data is the natural next "
                "input, not EPC data."
            )
            p_use.add_run(
                " so the output is a prioritisation signal, not a buy or don't-buy answer."
            )
        else:
            p_use.add_run(
                "This is what the tool is actually for: a buyer, landlord, or council retrofit "
                "scheme screens a shortlist of properties by postcode and gets back a ranked "
                "headroom probability, far cheaper than sending an assessor to every candidate "
                "address. It is not a valuation tool, and I want to be explicit about that "
                "rather than let the accuracy number imply more than it does. The two shaded "
                "rows in Table 2 are genuine model errors on real properties, not curated to "
                "look good, and neither purchase price nor the cost of the retrofit itself is "
                "estimated anywhere in this pipeline. A real investment decision, buying a "
                "specific Bristol property because it looks like a good retrofit candidate, "
                "still needs a separate estimate of retrofit cost against the value uplift it "
                "would actually produce, neither of which the EPC register or this model "
                "supplies; Land Registry Price Paid Data is the natural next input for that, not "
                "more EPC fields. What this pipeline gives a real user is a prioritisation "
                "signal, cheap to run at city scale, for where to look closer, not a buy or "
                "don't-buy answer by itself."
            )
    else:
        doc.add_paragraph(
            "[CASE STUDY: run src/bristol_case_study.py to generate real Bristol predictions "
            "from the held-out test set, then re-generate this report.]"
        )

    # ------------------------------------------------------------------
    # 8. Ethical Considerations
    # ------------------------------------------------------------------
    doc.add_heading("8. Ethical Considerations", level=1)
    if safe:
        p_eth = doc.add_paragraph()
        p_eth.add_run(
            "Published under the Open Government Licence with no directly identifying data. EPC "
            "records are property-addressable, so linkage with other datasets could re-identify "
            "a dwelling; no such linkage is performed. A false negative deprioritises a property "
            "that genuinely needs intervention; a false positive only wastes assessor time. That "
            "is why Section 6 recommends recall-oriented calibration: a person decides which "
            "flagged properties get a visit."
        )
        fm.add(p_eth,
            "The model inherits whatever bias sits in the EPC assessments it is trained on, so "
            "an assessor still checking each flagged property in person is a real safeguard here, "
            "not a formality."
        )
        p_eth.add_run(
            " The pipeline is public at https://github.com/KNHNF/epc-retrofit-potential-ml for "
            "independent checking."
        )
    else:
        doc.add_paragraph(
            "The dataset is published by the Ministry of Housing, Communities and Local Government "
            "under the Open Government Licence and contains no directly identifying personal data; "
            "however, EPC records are addressable to individual properties, so combining them with "
            "other publicly linkable datasets could in principle re-identify an occupant's dwelling. "
            "No attempt is made in this pipeline to link EPC records to any other individual-level "
            "dataset, and no property addresses are retained beyond the fields required for feature "
            "engineering."
        )
        doc.add_paragraph(
            "There is also a practical fairness consideration in how the model would be used. A false "
            "negative (a high-potential property scored as low-potential) means a property that "
            "genuinely warrants a retrofit intervention is deprioritised, a real cost to both the "
            "occupant and net-zero policy goals. A false positive wastes assessor time but causes no "
            "direct harm. This asymmetry is why Section 6 recommends recall-oriented threshold "
            "calibration rather than optimising for accuracy alone, and why the model is framed here "
            "as a prioritisation aid for human assessors rather than an automated decision-maker."
        )
        doc.add_paragraph(
            "The full pipeline, from data ingestion through to this report, is published at "
            "https://github.com/KNHNF/epc-retrofit-potential-ml so that the methodology, feature "
            "engineering decisions, and reported numbers can be independently checked rather than "
            "taken on trust."
        )

    # ------------------------------------------------------------------
    # 8. Conclusion
    # ------------------------------------------------------------------
    doc.add_heading("9. Conclusion", level=1)
    if winners:
        if safe:
            conclusion_lead = (
                f"{winners['best_auc_model']} led on every test-set metric, with the other "
                "tree-based model close behind."
            )
        else:
            conclusion_lead = (
                f"{winners['best_auc_model']} wins on every test-set metric, with the other "
                "tree-based model close behind and interpretable in its own right."
            )
    else:
        conclusion_lead = "[RESULT: state the best-performing model here once all notebooks have been run.]"
    if safe:
        p_conc = doc.add_paragraph()
        p_conc.add_run(
            "This paper presented a pipeline identifying high-retrofit-potential UK properties "
            f"from EPC open data. {conclusion_lead} Applied to real, held-out Bristol properties "
            "(Section 7), the model held its recall on an unseen city, not just an unseen time "
            "period. Current efficiency, CO2 intensity, construction age, and wall type drive the "
            "predictions, all physically interpretable. These numbers come with real caveats"
        )
        fm.add(p_conc,
            "EPC data-quality problems and a 200,000-record sample rather than the full 7.25 "
            "million eligible records (Section 6) mean this is a credible estimate, not a final one."
        )
        p_conc.add_run(
            ", set out in Section 6. Code and pipeline are public at "
            "https://github.com/KNHNF/epc-retrofit-potential-ml."
        )
    else:
        doc.add_paragraph(
            "I built this pipeline to predict retrofit headroom in UK homes from EPC data, not the "
            f"current energy rating most prior work predicts. {conclusion_lead} Section 7 pushed "
            "past the benchmark numbers and pointed the trained model at real, held-out Bristol "
            "properties; recall held up on a city the model had never specifically seen, which "
            "matters more to me than another decimal place on Table 1. Current efficiency score, "
            "CO2 intensity, construction age, and wall type drive the predictions, and all four "
            "line up with what retrofit policy already assumes. None of this is final: EPC "
            "data-quality problems and a 200,000-record sample rather than the full 7.25 million "
            "eligible records (Section 6) mean these are credible estimates, not the last word. "
            "Code and pipeline: "
            "https://github.com/KNHNF/epc-retrofit-potential-ml."
        )

    # ------------------------------------------------------------------
    # References. IEEE numeric style, listed in order of first citation.
    # ------------------------------------------------------------------
    doc.add_heading("References", level=1)

    add_reference(doc,
        "[1] Ministry of Housing, Communities and Local Government (2024) ",
        "Energy Performance of Buildings Data: England and Wales.",
        " Available from: https://epc.opendatacommunities.org [Accessed 9 July 2026]."
    )
    add_reference(doc,
        "[2] Seyedzadeh, S., Pour Rahimian, F., Glesk, I. and Roper, M. (2018) Machine learning "
        "for estimation of building energy consumption and performance: a review. ",
        "Visualization in Engineering.",
        " 6 (1), p. 5."
    )
    add_reference(doc,
        "[3] Pasichnyi, O., Wallin, J. and Kordas, O. (2019) Data-driven building archetypes for "
        "urban building energy modelling. ",
        "Energy.",
        " 181, pp. 360-377."
    )
    p_pedregosa = add_reference(doc, "[4] Pedregosa, F. ", "et al.", "")
    p_pedregosa.runs[1].italic = True
    p_pedregosa.add_run(" (2011) Scikit-learn: machine learning in Python. ").font.size = Pt(10)
    r = p_pedregosa.add_run("Journal of Machine Learning Research.")
    r.italic = True
    r.font.size = Pt(10)
    p_pedregosa.add_run(" 12, pp. 2825-2830.").font.size = Pt(10)
    add_reference(doc,
        "[5] Chen, T. and Guestrin, C. (2016) XGBoost: a scalable tree boosting system. ",
        "Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining.",
        " pp. 785-794."
    )
    add_reference(doc,
        "[6] Ng, A.Y. and Jordan, M.I. (2001) On discriminative vs. generative classifiers: a "
        "comparison of logistic regression and naive Bayes. ",
        "Advances in Neural Information Processing Systems 14 (NIPS 2001).",
        " pp. 841-848."
    )
    add_reference(doc,
        "[7] Breiman, L. (2001) Random forests. ",
        "Machine Learning.",
        " 45 (1), pp. 5-32."
    )
    add_reference(doc,
        "[8] Breiman, L., Friedman, J.H., Olshen, R.A. and Stone, C.J. (1984) ",
        "Classification and Regression Trees.",
        " Belmont, CA: Wadsworth."
    )
    add_reference(doc,
        "[9] Strobl, C., Boulesteix, A-L., Zeileis, A. and Hothorn, T. (2007) Bias in random "
        "forest variable importance measures: illustrations, sources and a solution. ",
        "BMC Bioinformatics.",
        " 8, p. 25."
    )
    add_reference(doc,
        "[10] Cortes, C. and Vapnik, V. (1995) Support-vector networks. ",
        "Machine Learning.",
        " 20 (3), pp. 273-297."
    )
    add_reference(doc,
        "[11] Beyer, K., Goldstein, J., Ramakrishnan, R. and Shaft, U. (1999) When is nearest "
        "neighbor meaningful? ",
        "International Conference on Database Theory (ICDT), Lecture Notes in Computer Science.",
        " Vol. 1540, pp. 217-235."
    )
    add_reference(doc,
        "[12] Varma, S. and Simon, R. (2006) Bias in error estimation when using cross-validation "
        "for model selection. ",
        "BMC Bioinformatics.",
        " 7, p. 91."
    )
    add_reference(doc,
        "[13] McNemar, Q. (1947) Note on the sampling error of the difference between correlated "
        "proportions or percentages. ",
        "Psychometrika.",
        " 12 (2), pp. 153-157."
    )
    add_reference(doc,
        "[14] Dietterich, T.G. (1998) Approximate statistical tests for comparing supervised "
        "classification learning algorithms. ",
        "Neural Computation.",
        " 10 (7), pp. 1895-1923."
    )
    add_reference(doc,
        "[15] Hardy, A. and Glew, D. (2019) An analysis of errors in the Energy Performance "
        "Certificate database. ",
        "Energy Policy.",
        " 129, pp. 1168-1178."
    )

    # Word count: Introduction through Conclusion only. Title, author block, abstract,
    # and references excluded (brief's 2000-word rule + standard academic convention).
    body_start = body_end = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "1. Introduction":
            body_start = i
        if p.text.strip() == "References":
            body_end = i
            break
    # Figure labels ("Fig. 1.") are captions, not prose, and single-column mode
    # already excludes them for free (they live inside a table cell, which
    # doc.paragraphs never sees). Two-column mode places the same label as a
    # top-level paragraph (add_floating_picture), so it WOULD get counted here
    # unless explicitly skipped, previously making the two layouts of the same
    # text disagree by a few words. Skipping the pattern in both modes keeps
    # the count comparable regardless of layout.
    fig_label_re = re.compile(r"^Fig\.\s*\d+\.?$")
    word_count = 0
    if body_start is not None and body_end is not None:
        for p in doc.paragraphs[body_start:body_end]:
            if p.style.name.startswith('Heading'):
                continue
            if fig_label_re.match(p.text.strip()):
                continue
            word_count += len(p.text.split())

    doc.add_paragraph()
    wc_para = doc.add_paragraph()
    wc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    limit_note = "2,100-word hard cap" if safe else "quality-first, not word-capped"
    wc_run = wc_para.add_run(
        f"Word count: {word_count} (Introduction to Conclusion, excluding title, "
        f"abstract, references, and footnotes, per standard academic convention). "
        f"Mode: {mode} ({limit_note})."
    )
    wc_run.italic = True
    wc_run.font.size = Pt(10)
    wc_run.font.name = 'Times New Roman'

    if two_column:
        # Endnotes default to lowercase roman numerals (i, ii, iii...);
        # match the arabic numbering the single-column doc's footnotes use.
        # Must be set on every section that exists by now (switch_columns
        # calls earlier added several), each section's endnotePr is its
        # own element, not inherited from an earlier section.
        for section in doc.sections:
            sectPr = section._sectPr
            endnotePr = OxmlElement('w:endnotePr')
            numFmt = OxmlElement('w:numFmt')
            numFmt.set(qn('w:val'), 'decimal')
            endnotePr.append(numFmt)
            sectPr.append(endnotePr)

    label = f"{mode}{'+2col' if two_column else ''}"
    doc.save(OUT_PATH)
    print(f"[{label}] Report saved to {OUT_PATH}")
    print(f"[{label}] Word count (Introduction to Conclusion, excl. title/abstract/references): {word_count}")
    return word_count


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else "both"
    if arg == "both":
        build_report("full")
        build_report("safe")
        build_report("safe", two_column=True)
    elif arg == "2col":
        build_report("safe", two_column=True)
    else:
        build_report(arg)
