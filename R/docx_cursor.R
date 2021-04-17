#' @export
#' @rdname cursor
#' @title set cursor in an rdocx object
#' @description a set of functions is available to manipulate
#' the position of a virtual cursor. This cursor will be used when
#' inserting, deleting or updating elements in the document.
#' @section cursor_begin:
#' Set the cursor at the beginning of the document, on the first element
#' of the document (usually a paragraph or a table).
#' @param x a docx device
#' @examples
#' library(officer)
#'
#' doc <- read_docx()
#' doc <- body_add_par(doc, "paragraph 1", style = "Normal")
#' doc <- body_add_par(doc, "paragraph 2", style = "Normal")
#' doc <- body_add_par(doc, "paragraph 3", style = "Normal")
#' doc <- body_add_par(doc, "paragraph 4", style = "Normal")
#' doc <- body_add_par(doc, "paragraph 5", style = "Normal")
#' doc <- body_add_par(doc, "paragraph 6", style = "Normal")
#' doc <- body_add_par(doc, "paragraph 7", style = "Normal")
#'
#' # default template contains only an empty paragraph
#' # Using cursor_begin and body_remove, we can delete it
#' doc <- cursor_begin(doc)
#' doc <- body_remove(doc)
#'
#' # Let add text at the beginning of the
#' # paragraph containing text "paragraph 4"
#' doc <- cursor_reach(doc, keyword = "paragraph 4")
#' doc <- slip_in_text(doc, "This is ", pos = "before", style = "Default Paragraph Font")
#'
#' # move the cursor forward and end a section
#' doc <- cursor_forward(doc)
#' doc <- body_add_par(doc, "The section stop here", style = "Normal")
#' doc <- body_end_section_landscape(doc)
#'
#' # move the cursor at the end of the document
#' doc <- cursor_end(doc)
#' doc <- body_add_par(doc, "The document ends now", style = "Normal")
#'
#' print(doc, target = tempfile(fileext = ".docx"))
cursor_begin <- function( x ){
  x$doc_obj$cursor_begin()
  x
}

#' @rdname cursor
#' @param id bookmark id
#' @section cursor_bookmark:
#' Set the cursor at a bookmark that has previously been set.
#' @export
#' @examples
#'
#' # cursor_bookmark ----
#'
#' doc <- read_docx()
#' doc <- body_add_par(doc, "centered text", style = "centered")
#' doc <- body_bookmark(doc, "text_to_replace")
#' doc <- body_add_par(doc, "A title", style = "heading 1")
#' doc <- body_add_par(doc, "Hello world!", style = "Normal")
#' doc <- cursor_bookmark(doc, "text_to_replace")
#' doc <- body_add_table(doc, value = iris, style = "table_template")
#'
#' print(doc, target = tempfile(fileext = ".docx"))
cursor_bookmark <- function( x, id ){
  x$doc_obj$cursor_bookmark(id)
  x
}

#' @export
#' @rdname cursor
#' @section cursor_end:
#' Set the cursor at the end of the document, on the last element
#' of the document.
cursor_end <- function( x ){
  x$doc_obj$cursor_end()
  x
}

#' @export
#' @rdname cursor
#' @param keyword keyword to look for as a regular expression
#' @section cursor_reach:
#' Set the cursor on the first element of the document
#' that contains text specified in argument \code{keyword}.
#' The argument \code{keyword} is a regexpr pattern.
cursor_reach <- function( x, keyword ){
  x$doc_obj$cursor_reach(keyword=keyword)
  x
}

#' @export
#' @rdname cursor
#' @section cursor_forward:
#' Move the cursor forward, it increments the cursor in the document.
cursor_forward <- function( x ){
  x$doc_obj$cursor_forward()
  x
}

#' @export
#' @rdname cursor
#' @section cursor_backward:
#' Move the cursor backward, it decrements the cursor in the document.
cursor_backward <- function( x ){
  x$doc_obj$cursor_backward()
  x
}
