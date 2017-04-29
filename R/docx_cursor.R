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
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_par("paragraph 1", style = "Normal") %>%
#'   body_add_par("paragraph 2", style = "Normal") %>%
#'   body_add_par("paragraph 3", style = "Normal") %>%
#'   body_add_par("paragraph 4", style = "Normal") %>%
#'   body_add_par("paragraph 5", style = "Normal") %>%
#'   body_add_par("paragraph 6", style = "Normal") %>%
#'   body_add_par("paragraph 7", style = "Normal") %>%
#'
#'   # default template contains only an empty paragraph
#'   # Using cursor_begin and body_remove, we can delete it
#'   cursor_begin() %>% body_remove() %>%
#'
#'   # Let add text at the beginning of the
#'   # paragraph containing text "paragraph 4"
#'   cursor_reach(keyword = "paragraph 4") %>%
#'   slip_in_text("This is ", pos = "before", style = "Default Paragraph Font") %>%
#'
#'   # move the cursor forward and end a section
#'   cursor_forward() %>%
#'   body_add_par("The section stop here", style = "Normal") %>%
#'   body_end_section(landscape = TRUE) %>%
#'
#'   # move the cursor at the end of the document
#'   cursor_end() %>%
#'   body_add_par("The document ends now", style = "Normal")
#'
#' print(doc, target = "cursor.docx")
#'
cursor_begin <- function( x ){
  x$doc_obj$cursor_begin()
  x
}

#' @rdname cursor
#' @param id bookmark id
#' @export
#' @examples
#'
#' # cursor_bookmark ----
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_par("centered text", style = "centered") %>%
#'   body_bookmark("text_to_replace") %>%
#'   body_add_par("A title", style = "heading 1") %>%
#'   body_add_par("Hello world!", style = "Normal") %>%
#'   cursor_bookmark("text_to_replace") %>%
#'   body_add_table(value = iris, style = "table_template")
#'
#' print(doc, target = "bookmark.docx")
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
#' @param keyword keyword to look for
#' @section cursor_reach:
#' Set the cursor on the first element of the document
#' that contains text specified in argument \code{keyword}.
#' @importFrom xml2 xml_find_first xml_path
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
