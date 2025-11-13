#' @export
#' @title  Write a 'Word' File
#' @description
#' `print.rdocx()` is the essential output function for creating Word files
#' with officer. It takes an `rdocx` object (created with [read_docx()] and
#' populated with content) and writes it to disk as a `.docx` file.
#'
#' This function performs all necessary post-processing operations before
#' writing the file.
#'
#' The function is typically called at the end of your document creation
#' workflow, after all content has been added with `body_add_*()` functions.
#'
#' @param x an `rdocx` object created with [read_docx()]
#' @param target path to the `.docx` file to write. The file will be created
#' or overwritten if it already exists. If `NULL` and `preview = FALSE`, the
#' function returns `NULL` without writing a file.
#' @param copy_header_refs,copy_footer_refs logical, default is FALSE.
#' If TRUE, copy the references to the header and footer in each section
#' of the body of the document. This parameter is experimental and may change
#' in a future version.
#' @param preview Save `x` to a temporary file and open it (default `FALSE`).
#' When `TRUE`, the document is saved to a temporary location and opened with
#' the system's default application for `.docx` files, useful for quick previewing
#' during development.
#' @param ... unused
#' @return The full path to the created `.docx` file (invisibly). This allows
#' chaining operations or capturing the output path for further use.
#' @examples
#' library(officer)
#'
#' # This example demonstrates how to create
#' # an small document -----
#'
#' ## Create a new Word document
#' doc <- read_docx()
#' doc <- body_add_par(doc, "hello world")
#' ## Save the document
#' output_file <- print(doc, target = tempfile(fileext = ".docx"))
#'
#' # preview mode: save to temp file and open locally ----
#' \dontrun{
#' print(doc, preview = TRUE)
#' }
#' @seealso Create a 'Word' document object with [read_docx()], add content with
#' functions [body_add_par()], [body_add_plot()],
#' [body_add_table()], change settings with [docx_set_settings()], set
#' properties with [set_doc_properties()],  read 'Word' styles with
#' [styles_info()].
print.rdocx <- function(x, target = NULL, copy_header_refs = FALSE,
                        copy_footer_refs = FALSE, preview = FALSE, ...) {
  if (preview) {
    file <- tempfile(fileext = ".docx")
    print.rdocx(x,
                target = file, copy_header_refs = copy_header_refs,
                copy_footer_refs = copy_footer_refs, preview = FALSE, ...
    )
    open_file(file)
    return(invisible(file))
  }
  if (is.null(target)) {
    cat("rdocx document with", length(x), "element(s)\n")
    cat("\n* styles:\n")

    style_names <- styles_info(x)
    style_sample <- style_names$style_type
    names(style_sample) <- style_names$style_name
    print(style_sample)
    return(invisible())
  }

  if (!grepl(x = target, pattern = "\\.(docx)$", ignore.case = TRUE)) {
    stop(target, " should have '.docx' extension.")
  }

  if (is_windows() && is_doc_open(target)) {
    stop(target, " is open. To write to this document, please, close it.")
  }

  # write the xml to a tempfile in a formatted way so that grepl is easy
  xml_str <- xml_document_to_chrs(x$doc_obj$get())
  xml_str <- process_sections_content(x, xml_str)
  xml_str <- process_footnotes_content(x$doc_obj, footnotes = x$footnotes, xml_str)
  xml_str <- process_comments_content(x$doc_obj, comments = x$comments, xml_str)
  xml_str <- convert_custom_styles_in_wml(xml_str, x$styles)
  xml_str <- fix_hyperlink_refs_in_wml(xml_str, x$doc_obj)
  xml_str <- fix_img_refs_in_wml(xml_str, x$doc_obj, x$doc_obj$relationship(), x$package_dir)
  xml_str <- fix_svg_refs_in_wml(xml_str, x$doc_obj, x$doc_obj$relationship(), x$package_dir)
  # make all id unique for document
  xml_str <- fix_empty_ids_in_wml(xml_str)
  if (copy_header_refs) {
    xml_str <- copy_header_references_everywhere(x, xml_str = xml_str)
  }
  if (copy_footer_refs) {
    xml_str <- copy_footer_references_everywhere(x, xml_str = xml_str)
  }

  x <- replace_xml_body_from_chr(x = x, xml_str = xml_str)


  process_docx_poured(
    doc_obj = x$doc_obj,
    relationships = x$doc_obj$relationship(),
    content_type = x$content_type,
    package_dir = x$package_dir
  )

  x$headers <- update_hf_list(part_list = x$headers, type = "header", package_dir = x$package_dir)
  x$footers <- update_hf_list(part_list = x$footers, type = "footer", package_dir = x$package_dir)

  for (header in x$headers) {
    xml_str <- xml_document_to_chrs(header$get())
    xml_str <- convert_custom_styles_in_wml(xml_str, x$styles)
    xml_str <- fix_empty_ids_in_wml(xml_str)
    xml_str <- fix_hyperlink_refs_in_wml(xml_str, header)
    xml_str <- fix_img_refs_in_wml(xml_str, header, header$relationship(), x$package_dir)
    xml_str <- fix_svg_refs_in_wml(xml_str, header, header$relationship(), x$package_dir)
    tf_xml <- tempfile(fileext = ".txt")
    writeLines(xml_str, tf_xml, useBytes = TRUE)
    header$replace_xml(tf_xml)
  }
  for (footer in x$footers) {
    xml_str <- xml_document_to_chrs(footer$get())
    xml_str <- convert_custom_styles_in_wml(xml_str, x$styles)
    xml_str <- fix_empty_ids_in_wml(xml_str)
    xml_str <- fix_hyperlink_refs_in_wml(xml_str, footer)
    xml_str <- fix_img_refs_in_wml(xml_str, footer, footer$relationship(), x$package_dir)
    xml_str <- fix_svg_refs_in_wml(xml_str, footer, footer$relationship(), x$package_dir)

    tf_xml <- tempfile(fileext = ".txt")
    writeLines(xml_str, tf_xml, useBytes = TRUE)
    footer$replace_xml(tf_xml)
  }
  if (TRUE) {
    xml_str <- xml_document_to_chrs(x$footnotes$get())
    xml_str <- convert_custom_styles_in_wml(xml_str, x$styles)
    xml_str <- fix_empty_ids_in_wml(xml_str)
    xml_str <- fix_hyperlink_refs_in_wml(xml_str, x$footnotes)
    xml_str <- fix_img_refs_in_wml(xml_str, x$footnotes, x$footnotes$relationship(), x$package_dir)
    xml_str <- fix_svg_refs_in_wml(xml_str, x$footnotes, x$footnotes$relationship(), x$package_dir)

    tf_xml <- tempfile(fileext = ".txt")
    writeLines(xml_str, tf_xml, useBytes = TRUE)
    x$footnotes$replace_xml(tf_xml)
  }
  if (TRUE) {
    xml_str <- xml_document_to_chrs(x$comments$get())
    xml_str <- convert_custom_styles_in_wml(xml_str, x$styles)
    xml_str <- fix_empty_ids_in_wml(xml_str)
    xml_str <- fix_hyperlink_refs_in_wml(xml_str, x$comments)
    xml_str <- fix_img_refs_in_wml(xml_str, x$comments, x$comments$relationship(), x$package_dir)
    xml_str <- fix_svg_refs_in_wml(xml_str, x$comments, x$comments$relationship(), x$package_dir)

    tf_xml <- tempfile(fileext = ".txt")
    writeLines(xml_str, tf_xml, useBytes = TRUE)
    x$comments$replace_xml(tf_xml)
  }

  int_id <- 1 # unique id identifier
  # make all id unique for footnote
  int_id <- correct_id(x$comments$get(), int_id)

  body <- xml_find_first(x$doc_obj$get(), "w:body")

  # If body is not ending with an sectPr, create a continuous one append it
  if (!xml_name(xml_child(body, search = xml_length(body))) %in% "sectPr") {
    str <- paste0(
      "<w:sectPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
      "<w:type w:val=\"continuous\"/></w:sectPr>"
    )
    xml_add_child(body, as_xml_document(str))
  }

  for (header in x$headers) {
    header$save()
  }
  for (footer in x$footers) {
    footer$save()
  }

  x$doc_obj$save()
  x$content_type$save()
  x$footnotes$save()
  x$comments$save()

  x$rel$write(file.path(x$package_dir, "_rels", ".rels"))
  write_docx_settings(x)

  # save doc properties
  if (nrow(x$doc_properties$data) > 0) {
    x$doc_properties["modified", "value"] <- format(Sys.time(), "%Y-%m-%dT%H:%M:%SZ")
    x$doc_properties["lastModifiedBy", "value"] <- Sys.getenv("USER")
    write_core_properties(x$doc_properties, x$package_dir)
  }
  if (nrow(x$doc_properties_custom$data) > 0) {
    write_custom_properties(x$doc_properties_custom, x$package_dir)
  }
  x <- sanitize_images(x, warn_user = FALSE)
  invisible(pack_folder(folder = x$package_dir, target = target))
}
