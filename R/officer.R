#' @title Manipulate Microsoft Word, PowerPoint and Excel Documents with 'officer'
#'
#' @description
#' The officer package facilitates access to and manipulation of
#' 'Microsoft Word', 'Microsoft PowerPoint' and 'Microsoft Excel' documents
#' from R. It also supports the writing of 'RTF' documents.
#'
#' Examples of usage are:
#'
#' * Create Word documents with tables, titles, TOC and graphics
#' * Importation of Word and PowerPoint files into data objects
#' * Write updated content back to a PowerPoint presentation
#' * Clinical reporting automation
#' * Production of reports from a shiny application
#'
#' To start with officer, read about [read_docx()], [read_pptx()],
#' [read_xlsx()] or [rtf_doc()].
#'
#' The package also provides several objects that can be printed
#' in 'R Markdown' documents for advanced Word or PowerPoint
#' reporting as [run_autonum()] and [block_caption()].
#' @section Word documents:
#'
#' Start with [read_docx()] and save with [print.rdocx()].
#'
#' Content is added with one function per kind of content, such as
#' [body_add_par()], [body_add_fpar()], [body_add_table()], [body_add_gg()]
#' and [body_add_toc()]. They add it at the cursor, which is moved with
#' [cursor_begin()], [cursor_reach()] and similar functions. For large documents,
#' see [body_append_start_context()].
#'
#' Sections are defined with [block_section()] and [body_end_block_section()].
#' Use [docx_summary()] to read the content of an existing document as a
#' data.frame.
#'
#' @section PowerPoint presentations:
#'
#' Start with [read_pptx()] and save with [print.rpptx()]. Add slides with
#' [add_slide()], then content with [ph_with()]. [layout_summary()] and
#' [plot_layout_properties()] show what a template provides.
#'
#' A new slide contains no shapes, and each call to `ph_with()` creates a new
#' one. Despite the `ph_` prefix, which stands for placeholder, it does not fill
#' an existing placeholder or replace what is already there. See section
#' *How it works* of [ph_with()].
#'
#' Use [pptx_summary()] to read the content of an existing presentation as a
#' data.frame.
#'
#' @section Excel workbooks:
#'
#' Start with [read_xlsx()] and save with [print.rxlsx()]. Add sheets with
#' [add_sheet()], then content with [sheet_write_data()] and
#' [sheet_add_drawing()].
#'
#' @section RTF documents:
#'
#' Start with [rtf_doc()] and save with [print.rtf()]. Add content with
#' [rtf_add()]. There is no cursor: elements can only be added one after the
#' other, at the end of the document, and an existing RTF file cannot be read.
#'
#' @section Formatted content:
#'
#' These objects work with several formats: [fpar()] and [ftext()] for formatted
#' text, [fp_text()] and [fp_par()] for their properties, [block_list()] for a
#' sequence of blocks.
#'
#' @section Copy an officer object:
#'
#' 'officer' objects of class `rdocx` or `rpptx` use R6 classes with reference
#' semantics. **Assignment does NOT create a copy**:
#'
#' ```
#' pptx1 <- read_pptx()
#' pptx2 <- pptx1  # pptx2 is a reference to pptx1, not a copy!
#' ```
#'
#' If you need independent documents (e.g., in loops), read the template
#' each time:
#'
#' ```
#' for (i in 1:10) {
#'   doc <- read_docx("template.docx")  # Read fresh each iteration
#'   # ... modify doc ...
#'   print(doc, target = paste0("output_", i, ".docx"))
#' }
#' ```
#'
#' @seealso The user documentation: \url{https://ardata-fr.github.io/officeverse/} and
#' manuals \url{https://davidgohel.github.io/officer/}
#' @docType package
#' @name officer
"_PACKAGE"
