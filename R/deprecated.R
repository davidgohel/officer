# old pptx chunk fun ----

#' @export
#' @title append text
#' @description append text in a placeholder.
#' The function let you add text to an existing
#' content in an exiisting shape, existing text will be preserved.
#' @note This function will be deprecated in a next release because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @section Usage:
#' If your goal is to add formatted text in a new shape, use \code{\link{ph_with}}
#' with a \code{\link{block_list}} instead of this function.
#' @inheritParams ph_remove
#' @param str text to add
#' @param style text style, a \code{\link{fp_text}} object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @param href hyperlink to reach when clicking the text
#' @param slide_index slide index to reach when clicking the text.
#' It will be ignored if \code{href} is not NULL.
ph_add_text <- function( x, str, type = "body", id = 1, id_chr = NULL, ph_label = NULL,
                         style = fp_text(font.size = 0), pos = "after",
                         href = NULL, slide_index = NULL ){
  .Deprecated(new = "ftext()")
  slide <- x$slide$get_slide(x$cursor)
  if( !is.null(id_chr)) {
    office_id <- id_chr
  } else {
    office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  }

  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )

  current_p <- xml_child(current_elt, "/a:p[last()]")
  if( inherits(current_p, "xml_missing") )
    stop("Could not find any paragraph in the selected shape.")

  r_shape_ <- to_pml(ftext(text = str, prop = style), add_ns = TRUE)

  if( pos == "after" ){
    runset <- xml_find_all(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]/a:p[last()]/a:r", office_id) )
    where_ <- length(runset)
  } else where_ <- 0

  new_node <- as_xml_document(r_shape_)

  if( !is.null(href)){
    slide$reference_hyperlink(href)
    rel_df <- slide$rel_df()
    id <- rel_df[rel_df$target == href, "id" ]

    apr <- xml_child(new_node, "a:rPr")
    str_ <- "<a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\"/>"
    str_ <- sprintf(str_, id)
    xml_add_child(apr, as_xml_document(str_) )
  } else if( !is.null(slide_index)){
    slide_name <- x$slide$names()[slide_index]
    slide$reference_slide(slide_name)
    rel_df <- slide$rel_df()
    id <- rel_df[rel_df$target == slide_name, "id" ]
    # add hlinkClick
    apr <- xml_child(new_node, "a:rPr")
    str_ <- "<a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\" action=\"ppaction://hlinksldjump\"/>"
    str_ <- sprintf(str_, id)
    xml_add_child(apr, as_xml_document(str_) )
  }


  xml_add_child(current_p, new_node, .where = where_ )

  slide$fortify_id()$save()

  x
}

#' @export
#' @title append paragraph
#' @description append a new empty paragraph in a placeholder.
#' The function let you add a new empty paragraph to an existing
#' content in an exiisting shape, existing paragraphs will be preserved.
#' @note This function will be deprecated in a next release because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @inheritParams ph_remove
#' @param level paragraph level
#' @inheritSection ph_add_text Usage
ph_add_par <- function( x, type = "body", id = 1, id_chr = NULL, level = 1, ph_label = NULL ){

  slide <- x$slide$get_slide(x$cursor)
  .Deprecated(new = "ftext()")

  if( !is.null(id_chr)) {
    office_id <- id_chr
  } else {
    office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  }
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )
  current_p <- xml_child(current_elt, "/p:txBody")

  if( inherits(current_p, "xml_missing") ){
    if( level > 1 )
      p_shape <- sprintf("<a:p><a:pPr lvl=\"%.0f\"/></a:p>", level - 1)
    else
      p_shape <- "<a:p/>"

    simple_shape <- paste0( "<p:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                            "<a:bodyPr/><a:lstStyle/>",
                            p_shape, "</p:txBody>")
    xml_add_child(current_elt, as_xml_document(simple_shape) )
  } else {
    if( level > 1 ){
      simple_shape <- sprintf(paste0( ap_ns_yes, "<a:pPr lvl=\"%.0f\"/></a:p>" ), level - 1)
    } else
      simple_shape <- paste0(ap_ns_yes, "</a:p>")
    xml_add_child(current_p, as_xml_document(simple_shape) )
  }
  x
}



#' @export
#' @title append fpar
#' @description append \code{fpar} (a formatted paragraph) in a placeholder
#' The function let you add a new formatted paragraph (\code{\link{fpar}})
#' to an existing content in an existing shape, existing paragraphs
#' will be preserved.
#' @note This function will be deprecated in a next release because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @inheritParams ph_remove
#' @param value fpar object
#' @param level paragraph level
#' @param par_default specify if the default paragraph formatting
#' should be used.
#' @inheritSection ph_add_text Usage
#' @seealso \code{\link{fpar}}
ph_add_fpar <- function( x, value, type = "body", id = 1, id_chr = NULL, ph_label = NULL,
                         level = 1, par_default = TRUE ){
  .Deprecated(new = "ftext()")
  slide <- x$slide$get_slide(x$cursor)

  if( !is.null(id_chr)) {
    office_id <- id_chr
  } else {
    office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  }
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )

  current_p <- xml_child(current_elt, "/p:txBody")

  node <- as_xml_document(to_pml(value, add_ns = TRUE))

  if( par_default ){
    # add default pPr
    ppr <- xml_child(node, "/a:pPr")
    empty_par <- as_xml_document(paste0("<a:pPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                                        "</a:pPr>"))
    xml_replace(ppr, empty_par )
  }
  ppr <- xml_child(node, "/a:pPr")
  if( level > 1 ){
    xml_attr(ppr, "lvl") <- sprintf("%.0f", level - 1)
  }
  if( inherits(current_p, "xml_missing") ){
    simple_shape <- paste0( "<p:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                            "<a:bodyPr/><a:lstStyle/></p:txBody>")
    newnode <- as_xml_document(simple_shape)
    xml_add_child(newnode, node)
    xml_add_child(current_elt, newnode )
  } else {
    xml_add_child(current_p, node )
  }

  x
}


# old docx chunk fun ----


#' @export
#' @title add a column break
#' @description add a column break into a Word document. A column break
#' is used to add a break in a multi columns section in a Word
#' Document.
#'
#' This function will be deprecated in the next release because it is not
#' efficient and make users write complex code, use [run_columnbreak()] instead.
#' @param x an rdocx object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
slip_in_column_break <- function( x, pos = "before" ){
  xml_elt <- paste0( wr_ns_yes, "<w:br w:type=\"column\"/>", "</w:r>")
  slip_in_xml(x = x, str = xml_elt, pos = pos)
}


#' @export
#' @title append seq field
#' @description append seq field into a paragraph of an rdocx object.
#' This feature is only available when document are edited with Word,
#' when edited with Libre Office or another program, seq field will not
#' be calculated and not displayed.
#'
#' This function will be deprecated in the next release because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @param x an rdocx object
#' @param str seq field value
#' @param style text style
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
slip_in_seqfield <- function( x, str, style = NULL, pos = "after" ){

  .Deprecated(new = "ftext()")
  if( is.null(style) )
    style <- x$default_styles$character

  style_id <- get_style_id(data = x$styles, style=style, type = "character")
  xml_elt_1 <- paste0(wr_ns_yes,
                      "<w:rPr/>",
                      "<w:fldChar w:fldCharType=\"begin\"/>",
                      "</w:r>")
  xml_elt_2 <- paste0(wr_ns_yes,
                      "<w:rPr/>",
                      sprintf("<w:instrText xml:space=\"preserve\">%s</w:instrText>", str ),
                      "</w:r>")
  xml_elt_3 <- paste0(wr_ns_yes,
                      "<w:rPr/>",
                      "<w:fldChar w:fldCharType=\"end\"/>",
                      "</w:r>")

  if( pos == "after"){
    slip_in_xml(x = x, str = xml_elt_1, pos = pos)
    slip_in_xml(x = x, str = xml_elt_2, pos = pos)
    slip_in_xml(x = x, str = xml_elt_3, pos = pos)
  } else {
    slip_in_xml(x = x, str = xml_elt_3, pos = pos)
    slip_in_xml(x = x, str = xml_elt_2, pos = pos)
    slip_in_xml(x = x, str = xml_elt_1, pos = pos)
  }

}

#' @export
#' @title append text
#' @description append text into a paragraph of an rdocx object.
#'
#' This function will be deprecated in the next release because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @param x an rdocx object
#' @param str text
#' @param style text style
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @param hyperlink turn the text into an external hyperlink
slip_in_text <- function( x, str, style = NULL, pos = "after", hyperlink = NULL ){

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
#' @title append an image
#' @description append an image into a paragraph of an rdocx object.
#'
#' This function will be deprecated in the next release because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @param x an rdocx object
#' @param src image filename, the basename of the file must not contain any blank.
#' @param style text style
#' @param width height in inches
#' @param height height in inches
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
slip_in_img <- function( x, src, style = NULL, width, height, pos = "after" ){

  .Deprecated(new = "ftext()")
  if( is.null(style) )
    style <- x$default_styles$character

  new_src <- tempfile( fileext = gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", src) )
  file.copy( src, to = new_src )

  style_id <- get_style_id(data = x$styles, style=style, type = "character")

  ext_img <- external_img(new_src, width = width, height = height)
  xml_elt <- to_wml(ext_img)
  xml_elt <- paste0(wp_ns_yes, "<w:pPr/>", xml_elt, "</w:p>")

  drawing_node <- xml_find_first(as_xml_document(xml_elt), "//w:r/w:drawing")

  wml_ <- paste0(wr_ns_yes, "<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>%s</w:r>")
  xml_elt <- sprintf(wml_, style_id, as.character(drawing_node) )

  slip_in_xml(x = x, str = xml_elt, pos = pos)
}


#' @export
#' @title add a wml string into a Word document
#' @description The function add a wml string into
#' the document after, before or on a cursor location.
#'
#' This function will be deprecated in the next release because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
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
