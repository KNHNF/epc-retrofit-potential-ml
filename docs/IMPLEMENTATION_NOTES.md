# Implementation notes

The reasoning behind a few non-obvious decisions in `src/`, pulled out of
long code comments so the code itself stays short. Read this if you want
the "why", the code just states the "what".

## make_table_images.py: booktabs table style

Published papers use "booktabs" style tables: a thick rule above the
header, a thin rule under it, a thick rule at the bottom, no vertical
lines and no lines between data rows. The full grid this project used
before read as a spreadsheet export, not a typeset table.

`Cell.visible_edges` looks like the right tool (set `'T'` on the header
row, `'B'` on the last row, nothing elsewhere), and it does draw the
right lines, but it has a real bug: it makes the whole cell invisible,
fill colour and text included, the moment a cell's edges are not the
full default set. Confirmed by testing a single cell on its own: all
four edges renders fine, one edge renders nothing at all.

The workaround: leave every cell's edges at their default (closed), but
coloured to match that cell's own fill so they don't show, then draw
the three real rules as plain line segments across the whole table,
independent of the cell grid. This needs the table's bounding box set
to the full axes (`[0, 0, 1, 1]`), so each row's vertical position is
known in advance (`1/n_rows`) rather than read from cell geometry that
is not finalised until the plot actually draws.

## generate_report.py: footnotes in single-column, endnotes in two-column

python-docx has no built-in support for footnotes or endnotes at all,
this builds the `word/footnotes.xml` / `word/endnotes.xml` document
part by hand, including the separator entries the OOXML schema
requires even when there are zero notes. Note text lives in that
separate part, not in `doc.paragraphs`, which is also why it never
counts toward the report's own word count.

Two-column mode uses endnotes instead of footnotes for a specific
reason: Word ties a footnote area's column count to the section body's
column count, so a footnote long enough to spill from column 1's
footnote area into column 2 forces Word to end the page early and
strand blank space below the shorter column, confirmed in real Word.
Endnotes collect once at the very end of the document instead of
per-page, so they never interact with column layout at all. Single-
column mode keeps ordinary footnotes, confirmed to render correctly
there since there is no column for them to spill across.

## make_table_images.py: why tables become pictures in the two-column build

A real Word table cannot float across both columns of a two-column
section the way a picture can. Tried three different ways to force it
(a section break, a text box, Word's own `w:tblpPr` table-positioning
property), all three failed in real Microsoft Word, checked by actually
opening the file each time, not assumed from the XML looking right.
Rendering the same data as an image sidesteps the limitation, since
pictures can float across columns and a table's data can be drawn to
look identical to the real Word table used in the single-column build.

## two_column_layout.py: display equations don't use a section break either

An earlier version switched to 1 column just for a display equation's
paragraph, then back to 2 after. Wrong in real Word (LibreOffice
rendered it fine and hid the problem): a continuous section cannot
change column count partway down a column, only at the top of a fresh
page, so Word blanked the rest of the current page next to a six-word
equation. Fixed the same way as floating figures, no section break,
the equation floats in a DrawingML text box instead. A first attempt
used the older VML text box format, which does not render OMML
equations inside it at all, silently, no error, only the newer
DrawingML format does.

## two_column_layout.py: a table cannot actually float across columns

Tried making a real python-docx table span both columns using Word's
own table-positioning property (`w:tblpPr`), the table equivalent of
the anchor that lets pictures float. It does not work: Word still
clips the table's visible width to whichever column it started in,
confirmed in real Word, not assumed from the XML. That is the reason
`make_table_images.py` exists at all, tables that need to span both
columns are rendered as images and floated as pictures instead.

## two_column_layout.py: compatibility mode and heading widow control

python-docx's default template ships with an old Word compatibility
setting that makes Word 2013+ open every generated file in
Compatibility Mode. Confirmed as a real problem, not cosmetic: a
floating equation box rendered as completely invisible until the file
was converted out of Compatibility Mode by hand. `set_compatibility_mode_full`
fixes the setting so nobody has to do that conversion manually.

Word's built-in Heading 1/2/3 styles also keep a heading on the same
page as the text after it, and keep the heading's own lines together.
On a tall single-column page there is usually enough room that this
never matters. In a narrow two-column layout there often is not: if a
heading and the next paragraph do not both fit in what is left of the
current column, Word pushes the whole thing to the top of the next
column and leaves the current one stranded, which turned out to be the
single biggest cause of blank space in the two-column build, more than
figures were. `disable_heading_widow_control` turns that behaviour off.

## two_column_layout.py: the font theme-leak bug

Setting `style.font.name` on python-docx's default styles does not
reliably work: those styles reference a theme font, and python-docx's
own file-saving code prefers the theme reference over the explicit one
you just set, only on save, not in memory (so it looks correct until
you actually open the file). `set_style_font` sets the font at a lower
level that bypasses the theme reference entirely.

## two_column_layout.py: deciding what floats versus what shrinks

Don't float everything, forcing a section break around every wide
figure leaves near-empty columns next to short ones, the opposite
whitespace problem. The rule used here, matching how real IEEE papers
do it: a figure with one or two simple series and short labels shrinks
to column width and flows inline, no column switch needed. Dense
content, correlation matrices, bar charts with long feature names,
side-by-side subplot pairs, or any table wider than about three
columns, floats at full page width instead. If a marker would need to
zoom in to read it at column width, it floats.

## two_column_layout.py: equations

Word's own equations are OMML (`m:oMath`), not something python-docx or
LaTeX renders directly. `add_equation()` converts a LaTeX string to
MathML via `latex2mathml`, then to OMML via Microsoft's own
MML2OMML.XSL stylesheet, which ships with every Office install. Real
typeset equations, not a screenshot pasted in as an image.

## two_column_layout.py: floating pictures and captions

`doc.add_picture()` only builds an inline drawing (`wp:inline`), which
is always constrained to the width of whatever column it sits in. To
span both columns of a two-column layout, this builds a `wp:anchor`
drawing instead, done by hand at the XML level since python-docx has no
built-in support for it.

Figure captions sit below the image, left-aligned, bold, not italic,
matching IEEE style (only table captions go above). An earlier version
put a bare tag above the image instead, changed for two reasons: a bare
tag next to a real caption elsewhere looks unfinished, and a caption
placed above a floating anchor is genuinely unreliable, the anchor's
vertical position is computed relative to the paragraph holding it, and
a multi-line caption's height is not known yet at that point.

No section break is used when floating a picture: an earlier attempt
that used one left stranded whitespace on the page, since Word inserts
a section break as a full paragraph with its own space.

## impact_and_fairness.py: matching the test set back to full records

The fairness and impact checks need columns (current energy rating,
heating cost, CO2) that are not in the model's own saved test outputs,
those only ever held the columns the model itself uses as inputs. This
loads `test_meta.csv` (the exact row order the model's predictions were
saved in) and matches it back to the full test parquet file by UPRN, so
the extra columns line up with the right property, not just the right
row number.

## walk_forward_evaluation.py: memory and compute constants

The full 7.25 million row training file does not fit in memory as one
pandas frame on a constrained machine, so each year is read straight
from the parquet file with a row-group filter (pyarrow skips anything
outside the date range before it becomes a DataFrame), then sampled
down immediately, per year rather than after combining, which keeps
memory bounded no matter how many years a fold covers. Not strictly
needed on Kaggle or Colab, where RAM can usually hold the whole file,
but there is no downside to leaving it as is.

## bootstrap_ci.py, walk_forward_evaluation.py: what the resampling proves

Random Forest's lead over XGBoost is small enough that it could be one
lucky test sample. Resampling the same test set with replacement 2,000
times and checking the gap holds across resamples is a standard way to
tell a real ordering from noise, more direct than reporting a single
p-value on its own.

The walk-forward check reuses the main model's already-tuned settings
rather than a fresh search per fold, on purpose: a full nested search
at every fold would take far longer than this project had time for. It
answers "does the ranking still work over time", not "what are the
best settings for each year".
