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
  x$cursor <- "/w:document/w:body/*[1]"
  x
}

#' @export
#' @rdname cursor
#' @section cursor_end:
#' Set the cursor at the end of the document, on the last element
#' of the document.
cursor_end <- function( x ){
  len <- length(x)
  if( len < 2 ) x$cursor <- "/w:document/w:body/*[1]"
  else x$cursor <- sprintf("/w:document/w:body/*[%.0f]", len - 1 )
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
  xpath_ <- sprintf("/w:document/w:body/*[contains(./w:r/w:t/text(),'%s')]", keyword)
  x$cursor <- xml_find_first(x$xml_doc, xpath_) %>% xml_path()
  x
}

#' @export
#' @rdname cursor
#' @section cursor_forward:
#' Move the cursor forward, it increments the cursor in the document.
cursor_forward <- function( x ){
  xpath_ <- paste0(x$cursor, "/following-sibling::*" )
  x$cursor <- xml_find_first(x$xml_doc, xpath_ ) %>% xml_path()
  x
}

#' @export
#' @rdname cursor
#' @section cursor_backward:
#' Move the cursor backward, it decrements the cursor in the document.
cursor_backward <- function( x ){
  xpath_ <- paste0(x$cursor, "/preceding-sibling::*" )
  x$cursor <- xml_find_first(x$xml_doc, xpath_ ) %>% xml_path()
  x
}

