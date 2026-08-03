"""
two_column_layout.py
Reusable helpers for IEEE-style two-column Word reports that actually look
like a published paper: no stranded whitespace, figures sized to fit their
column (or floating full width when they need to), bordered figure boxes,
and real typeset equations. Built for python-docx, which has no native
support for any of this.

Companion to the academic-docx skill. Copy this file into the report
project's src/ directory and import from it; do not import across projects.

Three things python-docx cannot do out of the box, all solved here:

1. Two columns. `switch_columns(doc, n)` inserts a continuous section break
   and sets the column count. Column changes are cheap; the real skill is
   knowing when NOT to use them (see below).

2. A figure that spans both columns. The naive approach is a continuous
   section break to 1 column around every wide figure. Don't do that: Word
   does not auto-balance a two-column section that is too short to fill
   both columns before a forced break, so every such figure stamps a
   near-empty column next to it. Real IEEE papers don't do this either.
   They float the image over the page, wrapped top-and-bottom, which lets
   it cross the column gutter without breaking the section at all.
   `add_floating_picture()` does this via a raw `wp:anchor` (python-docx's
   `add_picture` only builds inline `wp:inline` drawings, which are always
   constrained to the current column). This is the single highest-value
   technique in this file, use it for any figure/table that is too wide,
   dense, or label-heavy to survive being shrunk to column width.

3. Equations. Word's own equations are OMML (`m:oMath`), not something
   python-docx or LaTeX renders directly. `add_equation()` takes a LaTeX
   string, converts it to MathML via `latex2mathml` (pip install
   latex2mathml), then to OMML via Microsoft's own MML2OMML.XSL stylesheet,
   which ships with every Office install. Real typeset equations, not a
   screenshot pasted in as an image.

## Deciding what goes in-column vs floating

Don't float everything, that's what caused the original whitespace problem
in the other direction (every figure forcing a break). The rule that
matches how real IEEE papers do it:

- A figure with 1-2 simple series and short labels: shrink to column width
  (~2.7-2.9in on a 1in-margin A4 page) and place inline like a normal
  paragraph. No column switch needed at all, it just flows.
- A figure with dense content that would go illegible at column width
  (correlation matrices, permutation-importance bar charts with long
  feature-name labels, side-by-side subplot pairs like ROC+PR) or any
  table wider than ~3 columns: use `add_floating_picture()` at full page
  width instead of shrinking it.
- When in doubt, ask "would a marker need to zoom in to read this at
  column width?" If yes, float it.

## Usage sketch

    from two_column_layout import (
        switch_columns, add_floating_picture, add_bordered_picture,
        add_equation, set_page_a4,
    )

    doc = Document()
    set_page_a4(doc)                 # do this before anything else
    doc.add_paragraph("Title, authors, abstract: stay single column.")
    switch_columns(doc, 2)           # body starts here

    doc.add_paragraph("Two-column running text...")
    add_bordered_picture(doc, "fig_simple.png", "Fig. 1. Caption.", width_in=2.7)

    p = doc.add_paragraph()
    add_equation(p, r"F_1 = 2 \cdot \frac{P \cdot R}{P + R}")

    add_floating_picture(doc, "fig_dense_heatmap.png", "Fig. 2. Caption.",
                          width_in=6.2)   # crosses both columns, no break

Every function here was prototyped and visually verified (LibreOffice PDF
render, not just "the XML parses") before being written up. If you extend
this file, verify the same way, python-docx will happily save XML that
Word silently mangles on open (see the font bug note in add_style_font).
"""

import latex2mathml.converter
from lxml import etree

from docx.shared import Inches, Mm, Emu, Pt
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

# Path to Microsoft's own MathML-to-OMML stylesheet. Ships with every
# desktop Office install (Word, PowerPoint, Excel all use it). If this
# path doesn't exist on a given machine, search for MML2OMML.XSL under
# the Office install directory and update this constant.
MML2OMML_XSL_PATH = r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"


# ----------------------------------------------------------------------
# Page setup
# ----------------------------------------------------------------------

def set_page_a4(doc, margin_in=1.0):
    """python-docx's Document() defaults to US Letter. UK university
    submissions expect A4. Call this once, right after Document()."""
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Inches(margin_in)
    section.bottom_margin = Inches(margin_in)
    section.left_margin = Inches(margin_in)
    section.right_margin = Inches(margin_in)


def set_compatibility_mode_full(doc):
    """python-docx's bundled default template ships with
    `compatibilityMode` set to 14 (Word 2010) in settings.xml. Word 2013+
    then opens every generated document in "Compatibility Mode": the title
    bar says so, and it is not cosmetic. Confirmed directly in real Word:
    a DrawingML floating text box (used by `add_display_equation` for a
    two-column equation) rendered as nothing, completely invisible, no
    error, until the file was converted out of Compatibility Mode (File >
    Info > Convert). Call this once, right after `Document()`, so nobody
    has to do that by hand."""
    settings = doc.settings.element
    compat = settings.find(qn('w:compat'))
    if compat is None:
        compat = OxmlElement('w:compat')
        settings.append(compat)
    setting = None
    for el in compat.findall(qn('w:compatSetting')):
        if el.get(qn('w:name')) == 'compatibilityMode':
            setting = el
            break
    if setting is None:
        setting = OxmlElement('w:compatSetting')
        setting.set(qn('w:name'), 'compatibilityMode')
        setting.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word')
        compat.insert(0, setting)
    setting.set(qn('w:val'), '15')


def disable_heading_widow_control(doc, style_names=("Heading 1", "Heading 2", "Heading 3")):
    """Two-column layouts need this; single-column ones can take it or
    leave it. Word's built-in Heading 1/2/3 styles carry `keepNext` +
    `keepLines` (`w:pPr` widow/orphan control): a heading is kept on the
    same page as the paragraph after it, and its own lines are kept
    together. On a tall single-column page there's usually enough room
    that this never bites. In a narrow two-column layout there often
    isn't: if a heading and the start of its section don't both fit in
    the remaining space of the current column, Word bumps the whole
    thing to the top of the next column and leaves the rest of the
    current column blank. This is the single biggest cause of stranded
    whitespace in a two-column report that ISN'T about figures at all,
    confirmed by inspecting the actual style XML (`<w:keepNext/>
    <w:keepLines/>` sat right there in Heading1's `pPr`) and comparing a
    render where headings could fall wherever they wanted against one
    where they couldn't. Call this once, after `switch_columns(doc, 2)`,
    for a two-column document. Leaving it on for a single-column report
    is fine, there just isn't enough vertical pressure per page for it
    to matter."""
    for name in style_names:
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        pPr = style.element.find(qn('w:pPr'))
        if pPr is None:
            continue
        for tag in ('w:keepNext', 'w:keepLines'):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)


# ----------------------------------------------------------------------
# Fonts: the theme-leak bug
# ----------------------------------------------------------------------

def set_style_font(style, name, size=None, color=None):
    """Set a style's font properly, not just python-docx's `style.font.name`.

    The bug: python-docx's default template's styles (Heading 1/2, Normal,
    docDefaults) reference theme fonts (`w:asciiTheme="majorHAnsi"`, which
    resolves to Calibri). Setting `style.font.name = 'Times New Roman'`
    adds an explicit `w:ascii` attribute ALONGSIDE the theme one, but
    python-docx's own serializer favours the theme attribute once both are
    present, on save, not in memory. Verified by monkeypatching
    `Document.save` to dump the style XML immediately before
    serialization: correct in memory, wrong on disk, every time, for every
    style pulled from the built-in template. The fix is to strip the theme
    reference outright, not layer an explicit font on top of it.
    """
    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rpr.append(rFonts)
    _set_rfonts_explicit(rFonts, name)
    if size is not None:
        style.font.size = size
    if color is not None:
        style.font.color.rgb = color


def _set_rfonts_explicit(rFonts, name):
    for theme_attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
        attr_qn = qn('w:' + theme_attr)
        if rFonts.get(attr_qn) is not None:
            del rFonts.attrib[attr_qn]
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:cs'), name)


def fix_doc_defaults_font(doc, name):
    """The fallback every unstyled run inherits from (table cells under
    the built-in Table Grid style carry no font of their own, they fall
    through to this). Same theme-leak issue as styles, same fix."""
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
    _set_rfonts_explicit(rFonts, name)


# ----------------------------------------------------------------------
# Column switching
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
    """Continuous section break, no page break, just a column-count
    change from this point in the document onward. Note the OOXML
    quirk: a section's properties are stored in the paragraph that
    CLOSES that section, not the one that opens it, easy to get the
    before/after logic backwards when debugging."""
    section = doc.add_section(WD_SECTION_START.CONTINUOUS)
    set_section_columns(section, num_cols)
    return section


# ----------------------------------------------------------------------
# Floating (page-anchored) pictures: the whitespace fix
# ----------------------------------------------------------------------

_next_rel_height = [1]


def add_floating_picture(doc, image_path, label, width_in=6.2, caption_below=True,
                          caption_font="Times New Roman"):
    """A picture that floats over the page, spanning both columns of a
    two-column section, with body text wrapped top-and-bottom around it.
    No section break, so no stranded whitespace. Use this for any figure
    or table image too dense to survive being shrunk to one column.

    `doc.add_picture()` cannot do this: it only ever builds an inline
    drawing (`wp:inline`), which is always constrained to the width of
    whatever column it's placed in. This builds a `wp:anchor` instead.

    `label` is a real one-line descriptive caption ("Fig. 3. Pearson
    correlation matrix, numeric features and target."), placed BELOW the
    image, left-aligned, bold, not italic, matching standard IEEE style
    for figures (captions below; only table captions go above). Earlier
    drafts used a bare tag ("Fig. 1.") ABOVE the image, matching the
    92/100 NLP exemplar's box style, with the description living only in
    the surrounding body prose. Changed on two counts: a bare tag reads as
    unfinished next to a real caption, and a caption ABOVE a floating
    anchor is genuinely unreliable here, not just a style choice. The
    anchor's vertical position is computed relative to the paragraph that
    holds the drawing; when a real multi-line caption came first, its
    height was not always known yet by the time the anchor's position was
    resolved, and LibreOffice's renderer would visibly split the caption's
    own lines around the image (confirmed by rendering to PDF and
    inspecting the page, not assumed). Putting the caption after the
    image removes the ambiguity entirely, since nothing is anchored
    relative to it. Interpretation (what the figure means, why it
    matters) still belongs in body prose, not repeated in the caption.
    A real 1x1 bordered table cannot be used here (that is what
    `add_bordered_picture` does for column-width figures): a table inside
    a two-column section cannot cross the column gutter, confirmed broken
    three separate ways in real Word (see the table-floating notes
    below)."""
    rid, image = doc.part.get_or_add_image(image_path)
    width_emu = Emu(int(width_in * 914400))
    cx, cy = image.scaled_dimensions(width_emu, None)

    rel_height = _next_rel_height[0]
    _next_rel_height[0] += 1

    holder = doc.add_paragraph()
    holder.alignment = WD_ALIGN_PARAGRAPH.CENTER
    holder.paragraph_format.space_before = Pt(12)
    holder.paragraph_format.space_after = Pt(2)
    holder.paragraph_format.keep_with_next = True
    run = holder.add_run()

    drawing = OxmlElement('w:drawing')
    anchor = OxmlElement('wp:anchor')
    for k, v in {
        'behindDoc': '0', 'distT': '91440', 'distB': '457200', 'distL': '91440',
        'distR': '91440', 'simplePos': '0', 'locked': '0', 'layoutInCell': '1',
        'allowOverlap': '0', 'relativeHeight': str(rel_height),
    }.items():
        anchor.set(k, v)

    simplePos = OxmlElement('wp:simplePos')
    simplePos.set('x', '0')
    simplePos.set('y', '0')
    anchor.append(simplePos)

    posH = OxmlElement('wp:positionH')
    posH.set('relativeFrom', 'page')
    alignH = OxmlElement('wp:align')
    alignH.text = 'center'
    posH.append(alignH)
    anchor.append(posH)

    posV = OxmlElement('wp:positionV')
    posV.set('relativeFrom', 'paragraph')
    offV = OxmlElement('wp:posOffset')
    offV.text = '0'
    posV.append(offV)
    anchor.append(posV)

    extent = OxmlElement('wp:extent')
    extent.set('cx', str(cx))
    extent.set('cy', str(cy))
    anchor.append(extent)

    effectExtent = OxmlElement('wp:effectExtent')
    for k in ('l', 't', 'r', 'b'):
        effectExtent.set(k, '0')
    anchor.append(effectExtent)

    anchor.append(OxmlElement('wp:wrapTopAndBottom'))

    docPr = OxmlElement('wp:docPr')
    docPr.set('id', str(rel_height))
    docPr.set('name', f'Picture {rel_height}')
    anchor.append(docPr)

    cNvGraphicFramePr = OxmlElement('wp:cNvGraphicFramePr')
    graphicFrameLocks = OxmlElement('a:graphicFrameLocks')
    graphicFrameLocks.set('noChangeAspect', '1')
    cNvGraphicFramePr.append(graphicFrameLocks)
    anchor.append(cNvGraphicFramePr)

    graphic = OxmlElement('a:graphic')
    graphicData = OxmlElement('a:graphicData')
    graphicData.set('uri', PIC_URI)
    pic = OxmlElement('pic:pic')

    nvPicPr = OxmlElement('pic:nvPicPr')
    cNvPr = OxmlElement('pic:cNvPr')
    cNvPr.set('id', '0')
    cNvPr.set('name', image_path.split('/')[-1].split('\\')[-1])
    cNvPicPr = OxmlElement('pic:cNvPicPr')
    nvPicPr.append(cNvPr)
    nvPicPr.append(cNvPicPr)
    pic.append(nvPicPr)

    blipFill = OxmlElement('pic:blipFill')
    blip = OxmlElement('a:blip')
    blip.set(qn('r:embed'), rid)
    stretch = OxmlElement('a:stretch')
    stretch.append(OxmlElement('a:fillRect'))
    blipFill.append(blip)
    blipFill.append(stretch)
    pic.append(blipFill)

    spPr = OxmlElement('pic:spPr')
    xfrm = OxmlElement('a:xfrm')
    off = OxmlElement('a:off')
    off.set('x', '0')
    off.set('y', '0')
    ext = OxmlElement('a:ext')
    ext.set('cx', str(cx))
    ext.set('cy', str(cy))
    xfrm.append(off)
    xfrm.append(ext)
    prstGeom = OxmlElement('a:prstGeom')
    prstGeom.set('prst', 'rect')
    prstGeom.append(OxmlElement('a:avLst'))
    spPr.append(xfrm)
    spPr.append(prstGeom)
    spPr.append(_thin_border_line())
    pic.append(spPr)

    graphicData.append(pic)
    graphic.append(graphicData)
    anchor.append(graphic)
    drawing.append(anchor)
    run._r.append(drawing)

    lbl = doc.add_paragraph()
    lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lbl.paragraph_format.space_before = Pt(2)
    lbl.paragraph_format.space_after = Pt(14)
    lbl.paragraph_format.keep_together = True
    if label and label.strip():
        lbl_run = lbl.add_run(label)
        lbl_run.italic = True
        lbl_run.font.size = Pt(10)
        lbl_run.font.name = caption_font
    else:
        lbl.paragraph_format.space_after = Pt(8)

    return holder


def add_bordered_picture(doc, image_path, label, width_in=2.8,
                          caption_font="Times New Roman"):
    """In-column figure: centred image, italic caption below, glued together."""
    p_img = doc.add_paragraph()
    p_img.paragraph_format.space_before = Pt(12)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_img.add_run()
    run.add_picture(image_path, width=Inches(width_in))

    p_label = doc.add_paragraph()
    p_label.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_label.paragraph_format.space_before = Pt(0)
    p_label.paragraph_format.space_after = Pt(12)
    p_label.paragraph_format.keep_together = True
    if label and label.strip():
        r_label = p_label.add_run(label)
        r_label.italic = True
        r_label.font.size = Pt(10)
        r_label.font.name = caption_font
    return p_label


def _thin_border_line(color_hex="000000", width_emu=9525):
    """~0.75pt black outline for a picture shape (spPr/a:ln), matching a
    typeset paper's figure box, not a barely-visible hairline."""
    ln = OxmlElement('a:ln')
    ln.set('w', str(width_emu))
    solidFill = OxmlElement('a:solidFill')
    srgbClr = OxmlElement('a:srgbClr')
    srgbClr.set('val', color_hex)
    solidFill.append(srgbClr)
    ln.append(solidFill)
    return ln


# ----------------------------------------------------------------------
# Equations
# ----------------------------------------------------------------------

def latex_to_omml(latex, xsl_path=MML2OMML_XSL_PATH):
    """LaTeX -> MathML (latex2mathml) -> OMML (Microsoft's own XSLT,
    ships with every Office install). Real typeset Word equations, not
    a screenshot. Raises FileNotFoundError with a clear message if the
    stylesheet isn't at the expected path, search for MML2OMML.XSL
    under the local Office install directory and pass the real path in."""
    mathml = latex2mathml.converter.convert(latex)
    mathml_doc = etree.fromstring(mathml.encode("utf-8"))
    xslt = etree.parse(xsl_path)
    transform = etree.XSLT(xslt)
    omml_doc = transform(mathml_doc)
    return bytes(omml_doc)


def add_equation(paragraph, latex, xsl_path=MML2OMML_XSL_PATH):
    """Insert a real Word equation into `paragraph` from a LaTeX string.
    Centre the paragraph yourself if you want a display equation:
    `p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    add_equation(p, r"...")`."""
    omml_bytes = latex_to_omml(latex, xsl_path=xsl_path)
    omml_el = parse_xml(omml_bytes)
    paragraph._p.append(omml_el)
    return omml_el


def add_floating_table(doc, table, width_in=6.2):
    """DOES NOT achieve cross-column spanning, confirmed in real Word:
    `w:tblpPr` positions the table but Word still clips its visible
    rendering to the local column's width, unlike `wp:anchor` for
    pictures. Kept for reference and for cases where a table only needs
    modest same-column repositioning, not a genuinely wide table that
    must span both columns of a two-column section, for that, render the
    table as an image and float it with `add_floating_picture` instead
    (see the EPC retrofit coursework project's `make_table_images.py` and
    `generate_report.py` for the working pattern).

    Makes an already-built python-docx `Table` float over the page via
    native OOXML table positioning (`w:tblpPr`), so it spans both columns
    without a section break and without moving it into a text box.

    Two other approaches were tried and rejected, both confirmed broken in
    real Word, not just assumed:
    1. The original pattern, `switch_columns(doc, 1)` before the table and
       `switch_columns(doc, 2)` after: a continuous section break to fewer
       columns cannot start partway down a column, only at the top of a
       fresh page, so Word blanks the rest of the PREVIOUS page to get
       there, the identical bug `add_display_equation` had before its fix.
    2. Wrapping the table in a DrawingML floating text box (the same
       `wp:anchor`/`wps:wsp` mechanism that works for the equation): builds
       valid, schema-correct XML (confirmed by reading the raw docx XML,
       the table was genuinely there, correctly built) but Word renders it
       as nothing, no error. A `w:tbl` inside a `wps:txbx` is apparently
       not something Word's rendering engine supports, unlike OMML in the
       same container, which does render.

    `w:tblpPr` is the OOXML mechanism built specifically for a table that
    needs to float independent of the surrounding column/text layout,
    anchored to the page with text wrapping around it, the table-specific
    equivalent of `wp:anchor` for pictures. No section break, no text box,
    the table stays exactly where it already is in the document, this
    just adds positioning properties to its existing `tblPr`."""
    tbl = table._tbl
    tblPr = tbl.tblPr

    tblpPr = OxmlElement('w:tblpPr')
    tblpPr.set(qn('w:leftFromText'), '180')
    tblpPr.set(qn('w:rightFromText'), '180')
    tblpPr.set(qn('w:topFromText'), '120')
    tblpPr.set(qn('w:bottomFromText'), '120')
    tblpPr.set(qn('w:vertAnchor'), 'text')
    tblpPr.set(qn('w:horzAnchor'), 'page')
    tblpPr.set(qn('w:tblpXSpec'), 'center')
    tblpPr.set(qn('w:tblpY'), '1')
    tblPr.insert(0, tblpPr)

    width_twips = int(width_in * 1440)
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:type'), 'dxa')
    tblW.set(qn('w:w'), str(width_twips))
    return table


def add_display_equation(doc, latex, two_column=False, xsl_path=MML2OMML_XSL_PATH,
                          width_in=6.2):
    """A centred, full-width display equation. Use this instead of
    `add_equation` for anything beyond a short inline term.

    Earlier version of this function used a column-count section break
    (1 column for just this paragraph, back to 2 after). Confirmed in real
    Microsoft Word (not just LibreOffice, which rendered it fine and hid
    the problem) that this is wrong: switching a continuous section from 2
    columns to 1 mid-page forces Word to blank out the rest of the current
    page, because a full-width section cannot start partway down a column,
    it has to start at the top of a fresh page. That stranded almost an
    entire page next to a six-word sentence.

    Fixed the same way `add_floating_picture` fixes figures: no section
    break at all. First attempt floated a legacy VML `w:pict` text box;
    confirmed in real Word that OMML does not render inside a VML
    `v:textbox`, it renders as nothing, silently, no error. Switched to a
    modern DrawingML text box (`wps:wsp`/`wps:txbx`), the same `wp:anchor`
    float mechanism `add_floating_picture` uses, just with a text box
    instead of a `pic:pic` as the graphic payload. `a:spAutoFit` lets Word
    size the box to the equation's actual rendered size."""
    if not two_column:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_equation(p, latex, xsl_path=xsl_path)
        return p

    omml_bytes = latex_to_omml(latex, xsl_path=xsl_path)

    shape_id = _next_rel_height[0]
    _next_rel_height[0] += 1

    holder = doc.add_paragraph()
    holder.paragraph_format.space_before = Pt(6)
    holder.paragraph_format.space_after = Pt(6)
    run = holder.add_run()

    cx = Emu(int(width_in * 914400))
    cy = Emu(int(0.45 * 914400))
    drawing_xml = (
        '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wp:anchor behindDoc="0" distT="91440" distB="228600" distL="91440" '
        f'distR="91440" simplePos="0" locked="0" layoutInCell="1" allowOverlap="0" '
        f'relativeHeight="{shape_id}">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapTopAndBottom/>'
        f'<wp:docPr id="{shape_id}" name="EqBox{shape_id}"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp>'
        '<wps:cNvSpPr txBox="1"/>'
        '<wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/><a:ln><a:noFill/></a:ln>'
        '</wps:spPr>'
        '<wps:txbx><w:txbxContent>'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>'
        '</w:txbxContent></wps:txbx>'
        '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" '
        'rtlCol="0" anchor="ctr"><a:spAutoFit/></wps:bodyPr>'
        '</wps:wsp>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:anchor>'
        '</w:drawing>'
    )
    drawing_el = parse_xml(drawing_xml)
    inner_p = drawing_el.find('.//' + qn('w:txbxContent') + '/' + qn('w:p'))
    inner_p.append(parse_xml(omml_bytes))
    run._r.append(drawing_el)
    return holder
