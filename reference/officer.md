# Manipulate Microsoft Word, PowerPoint and Excel Documents with 'officer'

The officer package facilitates access to and manipulation of 'Microsoft
Word', 'Microsoft PowerPoint' and 'Microsoft Excel' documents from R. It
also supports the writing of 'RTF' documents.

Examples of usage are:

- Create Word documents with tables, titles, TOC and graphics

- Importation of Word and PowerPoint files into data objects

- Write updated content back to a PowerPoint presentation

- Clinical reporting automation

- Production of reports from a shiny application

To start with officer, read about
[`read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.md),
[`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md),
[`read_xlsx()`](https://davidgohel.github.io/officer/reference/read_xlsx.md)
or
[`rtf_doc()`](https://davidgohel.github.io/officer/reference/rtf_doc.md).

The package also provides several objects that can be printed in 'R
Markdown' documents for advanced Word or PowerPoint reporting as
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md)
and
[`block_caption()`](https://davidgohel.github.io/officer/reference/block_caption.md).

## Word documents

Start with
[`read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.md)
and save with
[`print.rdocx()`](https://davidgohel.github.io/officer/reference/print.rdocx.md).

Content is added with one function per kind of content, such as
[`body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/reference/body_add_fpar.md),
[`body_add_table()`](https://davidgohel.github.io/officer/reference/body_add_table.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/reference/body_add_gg.md)
and
[`body_add_toc()`](https://davidgohel.github.io/officer/reference/body_add_toc.md).
They add it at the cursor, which is moved with
[`cursor_begin()`](https://davidgohel.github.io/officer/reference/cursor.md),
[`cursor_reach()`](https://davidgohel.github.io/officer/reference/cursor.md)
and similar functions. For large documents, see
[`body_append_start_context()`](https://davidgohel.github.io/officer/reference/body_append_context.md).

Sections are defined with
[`block_section()`](https://davidgohel.github.io/officer/reference/block_section.md)
and
[`body_end_block_section()`](https://davidgohel.github.io/officer/reference/body_end_block_section.md).
Use
[`docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.md)
to read the content of an existing document as a data.frame.

## PowerPoint presentations

Start with
[`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md)
and save with
[`print.rpptx()`](https://davidgohel.github.io/officer/reference/print.rpptx.md).
Add slides with
[`add_slide()`](https://davidgohel.github.io/officer/reference/add_slide.md),
then content with
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md).
[`layout_summary()`](https://davidgohel.github.io/officer/reference/layout_summary.md)
and
[`plot_layout_properties()`](https://davidgohel.github.io/officer/reference/plot_layout_properties.md)
show what a template provides.

A new slide contains no shapes, and each call to
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md)
creates a new one. Despite the `ph_` prefix, which stands for
placeholder, it does not fill an existing placeholder or replace what is
already there. See section *How it works* of
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md).

Use
[`pptx_summary()`](https://davidgohel.github.io/officer/reference/pptx_summary.md)
to read the content of an existing presentation as a data.frame.

## Excel workbooks

Start with
[`read_xlsx()`](https://davidgohel.github.io/officer/reference/read_xlsx.md)
and save with
[`print.rxlsx()`](https://davidgohel.github.io/officer/reference/read_xlsx.md).
Add sheets with
[`add_sheet()`](https://davidgohel.github.io/officer/reference/add_sheet.md),
then content with
[`sheet_write_data()`](https://davidgohel.github.io/officer/reference/sheet_write_data.md)
and
[`sheet_add_drawing()`](https://davidgohel.github.io/officer/reference/sheet_add_drawing.md).

## RTF documents

Start with
[`rtf_doc()`](https://davidgohel.github.io/officer/reference/rtf_doc.md)
and save with
[`print.rtf()`](https://davidgohel.github.io/officer/reference/print.rtf.md).
Add content with
[`rtf_add()`](https://davidgohel.github.io/officer/reference/rtf_add.md).
There is no cursor: elements can only be added one after the other, at
the end of the document, and an existing RTF file cannot be read.

## Formatted content

These objects work with several formats:
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md) and
[`ftext()`](https://davidgohel.github.io/officer/reference/ftext.md) for
formatted text,
[`fp_text()`](https://davidgohel.github.io/officer/reference/fp_text.md)
and
[`fp_par()`](https://davidgohel.github.io/officer/reference/fp_par.md)
for their properties,
[`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md)
for a sequence of blocks.

## Copy an officer object

'officer' objects of class `rdocx` or `rpptx` use R6 classes with
reference semantics. **Assignment does NOT create a copy**:

    pptx1 <- read_pptx()
    pptx2 <- pptx1  # pptx2 is a reference to pptx1, not a copy!

If you need independent documents (e.g., in loops), read the template
each time:

    for (i in 1:10) {
      doc <- read_docx("template.docx")  # Read fresh each iteration
      # ... modify doc ...
      print(doc, target = paste0("output_", i, ".docx"))
    }

## See also

The user documentation: <https://ardata-fr.github.io/officeverse/> and
manuals <https://davidgohel.github.io/officer/>

## Author

**Maintainer**: David Gohel <david.gohel@ardata.fr>

Authors:

- David Gohel <david.gohel@ardata.fr>

- Stefan Moog <moogs@gmx.de>

- Mark Heckmann <heckmann.mark@gmail.com>
  ([ORCID](https://orcid.org/0000-0002-0736-7417))

Other contributors:

- ArData \[copyright holder\]

- Frank Hangler <frank@plotandscatter.com> (function
  body_replace_all_text) \[contributor\]

- Liz Sander <lsander@civisanalytics.com> (several documentation fixes)
  \[contributor\]

- Anton Victorson <anton@victorson.se> (fixes xml structures)
  \[contributor\]

- Jon Calder <jonmcalder@gmail.com> (update vignettes) \[contributor\]

- John Harrold <john.m.harrold@gmail.com> (function annotate_base)
  \[contributor\]

- John Muschelli <muschellij2@gmail.com> (google doc compatibility)
  \[contributor\]

- Bill Denney <wdenney@humanpredictions.com>
  ([ORCID](https://orcid.org/0000-0002-5759-428X)) (function
  as.matrix.rpptx) \[contributor\]

- Nikolai Beck <beck.nikolai@gmail.com> (set speaker notes for .pptx
  documents) \[contributor\]

- Greg Leleu <gregoire.leleu@gmail.com> (fields functionality in ppt)
  \[contributor\]

- Majid Eismann \[contributor\]

- Wahiduzzaman Khan (vectorization of remove_slide) \[contributor\]

- Hongyuan Jia <hongyuanjia@cqust.edu.cn>
  ([ORCID](https://orcid.org/0000-0002-0075-8183)) \[contributor\]

- Michael Stackhouse <mike.stackhouse@atorusresearch.com>
  \[contributor\]
