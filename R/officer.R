#' @title Manipulate Microsoft Word and PowerPoint Documents with 'officer'
#'
#' @description
#' The officer package facilitates access to and manipulation of
#' 'Microsoft Word' and 'Microsoft PowerPoint' documents from R.
#' It also supports the writing of 'RTF' documents.
#'
#' Examples of usage are:
#'
#' * Create Word documents with tables, titles, TOC and graphics
#' * Importation of Word and PowerPoint files into data objects
#' * Write updated content back to a PowerPoint presentation
#' * Clinical reporting automation
#' * Production of reports from a shiny application
#'
#' To start with officer, read about [read_docx()], [read_pptx()]
#' or [rtf_doc()].
#'
#' The package is also providing several objects that can be printed
#' in 'R Markdown' documents for advanced Word or PowerPoint
#' reporting as [run_autonum()] and [block_caption()].
#' @section Get Word content in a data.frame:
#'
#' While officer allows you to generate Word and PowerPoint documents,
#' an important feature is also the ability to read the content of existing
#' Word documents. Use [docx_summary()] to extract document content as a
#' structured data.frame, making it easy to analyze and process Word files
#' programmatically.
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
