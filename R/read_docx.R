#' @export
#' @title Create a 'Word' document object
#' @description `read_docx()` is the starting point for creating Word documents
#' from R. It creates an R object representing a Word document that can be
#' manipulated programmatically. When called without arguments, it creates an
#' empty document based on a default template. When provided with a path, it
#' reads an existing Word document (`.docx`) or template (`.dotx`) file.
#'
#' Once created, you can:
#' - Add content from R: Insert text, formatted paragraphs ([fpar()]),
#'   tables ([body_add_table()]), plots ([body_add_plot()]), images,
#'   page breaks, table of contents, and more using the `body_add_*()` family
#'   of functions.
#' - Read and inspect content: Use [docx_summary()] to extract and analyze
#'   the document's content, structure, and formatting as a data frame.
#' - Write to file: Save the document to disk using
#'   `print(x, target = "path/to/file.docx")`.
#'
#' @param path path to the docx file to use as base document.
#' `dotx` file are supported.
#' @return an object of class `rdocx`.
#' @section styles:
#'
#' The template file (specified via `path` or the default template) determines
#' the available paragraph styles, character styles, and table styles in your
#' document. These styles control the appearance of headings, body text, tables,
#' and other elements.
#'
#' When you use functions like `body_add_par(style = "heading 2")`, the style
#' name must exist in the template. You can:
#' - Use [styles_info()] to list all available styles in your document
#' - Create a custom template in Word with your organization's styles and branding
#' - Use the default template for standard documents
#'
#' The document layout (page size, margins, headers and footer content,
#' orientation) also comes from the
#' template and can be modified using [body_set_default_section()] or by
#' adding section breaks with [body_end_section_continuous()] and related functions.
#' @examples
#' library(officer)
#'
#' # This example demonstrates how to create
#' # an empty document -----
#'
#' ## Create a new Word document
#' doc <- read_docx()
#' ## Save the document
#' output_file <- print(doc, target = tempfile(fileext = ".docx"))
#'
#' @example inst/examples/example_read_docx_1.R
#' @example inst/examples/example_read_docx_2.R
#' @seealso Save a 'Word' document object to a file with [print.rdocx()],
#' add content with functions [body_add_par()], [body_add_plot()],
#' [body_add_table()], change settings with [docx_set_settings()], set
#' properties with [set_doc_properties()],  read 'Word' styles with
#' [styles_info()].
#' @section Illustrations:
#'
#' \if{html}{\figure{read_docx.png}{options: style="width:80\%;"}}
read_docx <- function(path = NULL) {
  if (!is.null(path) && !file.exists(path)) {
    stop("could not find file ", shQuote(path), call. = FALSE)
  }

  if (is.null(path)) {
    path <- system.file(package = "officer", "template/template.docx")
  }

  if (!grepl("\\.(docx|dotx)$", path, ignore.case = TRUE)) {
    stop("read_docx only support docx files", call. = FALSE)
  }

  package_dir <- tempfile()
  unpack_folder(file = path, folder = package_dir)

  obj <- structure(list(package_dir = package_dir),
                   .Names = c("package_dir"),
                   class = "rdocx"
  )

  obj$settings <- update_docx_settings_from_file(
    x = docx_settings(),
    file = file.path(package_dir, "word", "settings.xml")
  )

  obj$rel <- relationship$new()
  obj$rel$feed_from_xml(file.path(package_dir, "_rels", ".rels"))

  obj$doc_properties_custom <- read_custom_properties(package_dir)
  obj$doc_properties <- read_core_properties(package_dir)
  obj$content_type <- content_type$new(package_dir)
  obj$doc_obj <- body_part$new(
    package_dir,
    main_file = "document.xml",
    cursor = "/w:document/w:body/*[1]",
    body_xpath = "/w:document/w:body"
  )
  obj$styles <- read_docx_styles(package_dir)
  obj$officer_cursor <- officer_cursor(obj$doc_obj$get())

  obj$headers <- update_hf_list(part_list = list(), type = "header", package_dir = package_dir)
  obj$footers <- update_hf_list(part_list = list(), type = "footer", package_dir = package_dir)

  if (!file.exists(file.path(package_dir, "word", "comments.xml"))) {
    file.copy(
      system.file(package = "officer", "template", "comments.xml"),
      file.path(package_dir, "word", "comments.xml"),
      copy.mode = FALSE
    )
    obj$content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", "/word/comments.xml")
    )
  }

  obj$comments <- docx_part$new(
    package_dir,
    main_file = "comments.xml",
    cursor = "/w:comments/*[last()]", body_xpath = "/w:comments"
  )

  if (!file.exists(file.path(package_dir, "word", "footnotes.xml"))) {
    file.copy(
      system.file(package = "officer", "template", "footnotes.xml"),
      file.path(package_dir, "word", "footnotes.xml"),
      copy.mode = FALSE
    )
    obj$content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "/word/footnotes.xml")
    )
  }

  obj$footnotes <- footnotes_part$new(
    package_dir,
    main_file = "footnotes.xml",
    cursor = "/w:footnotes/*[last()]",
    body_xpath = "/w:footnotes"
  )

  default_refs <- obj$styles[obj$styles$is_default, ]
  obj$default_styles <- setNames(as.list(default_refs$style_name), default_refs$style_type)

  last_sect <- xml_find_first(obj$doc_obj$get(), "/w:document/w:body/w:sectPr[last()]")
  obj$sect_dim <- section_dimensions(last_sect)

  obj <- cursor_end(obj)
  obj
}


#' @export
#' @title Body xml document
#' @description Get the body document as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_body_xml(doc)
#' @keywords internal
docx_body_xml <- function( x ){
  x$doc_obj$get()
}
#' @export
#' @title xml element on which cursor is
#' @description Get the current block element as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer. If the
#' document is empty, this block will be set to NULL.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_current_block_xml(doc)
#' @keywords internal
docx_current_block_xml <- function( x ){
  ooxml_on_cursor(x$officer_cursor, x$doc_obj$get())
}

#' @export
#' @title Body xml document
#' @description Get the body document as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_body_relationship(doc)
#' @keywords internal
docx_body_relationship <- function( x ){
  x$doc_obj$relationship()
}
