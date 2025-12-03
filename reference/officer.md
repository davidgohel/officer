# Manipulate Microsoft Word and PowerPoint Documents with 'officer'

The officer package facilitates access to and manipulation of 'Microsoft
Word' and 'Microsoft PowerPoint' documents from R. It also supports the
writing of 'RTF' documents.

Examples of usage are:

- Create Word documents with tables, titles, TOC and graphics

- Importation of Word and PowerPoint files into data objects

- Write updated content back to a PowerPoint presentation

- Clinical reporting automation

- Production of reports from a shiny application

To start with officer, read about
[`read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.md),
[`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md)
or
[`rtf_doc()`](https://davidgohel.github.io/officer/reference/rtf_doc.md).

The package is also providing several objects that can be printed in 'R
Markdown' documents for advanced Word or PowerPoint reporting as
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md)
and
[`block_caption()`](https://davidgohel.github.io/officer/reference/block_caption.md).

## Get Word content in a data.frame

While officer allows you to generate Word and PowerPoint documents, an
important feature is also the ability to read the content of existing
Word documents. Use
[`docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.md)
to extract document content as a structured data.frame, making it easy
to analyze and process Word files programmatically.

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

- Hongyuan Jia <hongyuanjia@cqust.edu.cn>
  ([ORCID](https://orcid.org/0000-0002-0075-8183)) \[contributor\]

- Michael Stackhouse <mike.stackhouse@atorusresearch.com>
  \[contributor\]
