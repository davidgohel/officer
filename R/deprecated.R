#' @export
#' @title add a column break
#' @description add a column break into a Word document. A column break
#' is used to add a break in a multi columns section in a Word
#' Document.
#'
#' This function is deprecated and will be marked as defunct in
#' the next release because it is not efficient and
#' make users write complex code, use [run_columnbreak()]
#' instead.
#' @param x an rdocx object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
slip_in_column_break <- function( x, pos = "before" ){
  .Deprecated("slip_in_xml()")
  xml_elt <- paste0( wr_ns_yes, "<w:br w:type=\"column\"/>", "</w:r>")
  slip_in_xml(x = x, str = xml_elt, pos = pos)
}



#' @export
#' @title append text
#' @description append text into a paragraph of an rdocx object.
#'
#' This function is deprecated and will be marked as defunct
#' in the next release because it is not
#' efficient and make users write complex code.
#' Use instead [fpar()] to build
#' formatted paragraphs.
#' @param x an rdocx object
#' @param str text
#' @param style text style
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @param hyperlink turn the text into an external hyperlink
slip_in_text <- function( x, str, style = NULL, pos = "after", hyperlink = NULL ){
  .Deprecated("slip_in_xml()")
  if( is.null(style) )
    style <- x$default_styles$character

  style_id <- get_style_id(data = x$styles, style=style, type = "character")

  if( is.null(hyperlink) ) {
    xml_elt <- paste0( wr_ns_yes,
                       "<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>",
                       "<w:t xml:space=\"preserve\">%s</w:t></w:r>")
    xml_elt <- sprintf(xml_elt, style_id, htmlEscapeCopy(str))
  } else {
    hyperlink_id <- paste0("rId", x$doc_obj$relationship()$get_next_id())
    x$doc_obj$relationship()$add(
      id = hyperlink_id,
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target = hyperlink,
      target_mode = "External" )

    xml_elt <- paste0( "<w:hyperlink r:id=\"%s\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
                       "<w:r><w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>",
                       "<w:t xml:space=\"preserve\">%s</w:t></w:r></w:hyperlink>")
    xml_elt <- sprintf(xml_elt, hyperlink_id, style_id, htmlEscapeCopy(str))
  }

  slip_in_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add a wml string into a Word document
#' @description The function add a wml string into
#' the document after, before or on a cursor location.
#'
#' This function is deprecated and will be marked as defunct
#' in the next release because it is not efficient and make
#' users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @param x an rdocx object
#' @param str a wml string
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @keywords internal
slip_in_xml <- function(x, str, pos){

  xml_elt <- as_xml_document(str)
  pos_list <- c("after", "before")

  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  cursor_elt <- x$doc_obj$get_at_cursor()
  pos <- ifelse(pos=="after", length(xml_children(cursor_elt)), 1)
  xml_add_child(.x = cursor_elt, .value = xml_elt, .where = pos )
  x
}
