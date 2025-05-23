#' @export
#' @title Fast Append context to a Word document
#' @description
#' This function is used to append content to a Word document
#' in a fast way.
#'
#' It does not use the XML tree of the document
#' neither the cursor that is responsible for increasing the
#' performance of Word document generation when looping
#' over a large number of elements.
#'
#' This function must be used with the `write_elements_to_context()`
#' and `body_append_stop_context()` functions:
#'
#' 1. `body_append_start_context()` creates a context and returns
#'    a list with the context and the file connection.
#' 2. `write_elements_to_context()` writes the elements to the context
#'   file connection.
#' 3. `body_append_stop_context()` closes the file connection and
#'   replaces the XML in the document with the new XML.
#' @param x an rdocx object
#' @return `body_append_start_context()` returns a list representing the context
#' that contains:
#' - `doc`: the original document object
#' - `file_con`: the file connection to the context
#' - `file_path`: the path to the context file
#' - `final_str`: the final XML string to be appended to the document
#' later when calling `body_append_stop_context()`.
#'
#' This object should not be modified by the user but instead
#' passed to `write_elements_to_context()` and `body_append_stop_context()`.
#'
#' `write_elements_to_context()` returns the context object.
#'
#' `body_append_stop_context()` returns the `rdocx` object with the
#' cursor position set to the end of the document.
#' @importFrom utils tail
#' @rdname body_append_context
#' @example examples/body_append_context.R
#' @family functions for adding content
body_append_start_context <- function(x) {
  current_xml_file <- tempfile(fileext = ".txt")
  xml_doc <- docx_body_xml(x)
  current_xml_str <- as.character(xml_doc)
  writeLines(current_xml_str, current_xml_file)
  current_xml_str_lines <- readLines(current_xml_file)
  unlink(current_xml_file)

  section_start <- grepl("<w:sectPr", current_xml_str_lines)
  default_section_pos <- tail(which(section_start), n = 1)

  output_file <- tempfile(fileext = ".txt")
  con <- file(output_file, open = "w", encoding = "UTF-8")
  cat(
    current_xml_str_lines[seq_len(default_section_pos - 1)],
    sep = "\n",
    file = con,
    append = TRUE
  )

  out <- list(
    doc = x,
    file_con = con,
    file_path = output_file,
    final_str = current_xml_str_lines[
      default_section_pos:length(current_xml_str_lines)
    ]
  )
  class(out) <- "rdocx_append_context"
  out
}

#' @rdname body_append_context
#' @param ... elements to be written to the context. These can be
#' paragraphs, tables, images, etc. The elements should have an associated
#' `to_wml()` method that converts them to WML format.
#' @param context the context object created by `body_append_start_context()`.
#' @export
write_elements_to_context <- function(context, ...) {
  objects <- list(...)
  for (obj in objects) {
    cat(to_wml(obj), sep = "\n", file = context$file_con, append = TRUE)
  }
  context
}

#' @rdname body_append_context
#' @export
body_append_stop_context <- function(context) {
  cat(context$final_str, sep = "\n", file = context$file_con, append = TRUE)
  close(context$file_con)
  context$doc$doc_obj$replace_xml(context$file_path)
  unlink(context$file_path)
  x <- context$doc
  x$officer_cursor <- officer_cursor(x$doc_obj$get())
  x
}
