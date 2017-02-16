#' @export
#' @rdname cursor
#' @title set cursor in a docx object
#' @description a set of functions is available to manipulate
#' the position of a virtual cursor. This cursor will be used when
#' inserting, deleting or updating elements in the document.
#' @section cursor_begin:
#' Set the cursor at the beginning of the document, on the first element
#' of the document (usually a paragraph or a table).
#' @param x a docx device
cursor_begin <- function( x ){
  x$doc_obj$cursor_begin()
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

