#' @export
#' @title add page break
#' @description add a page break into an rdocx object
#' @param x an rdocx object
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' library(magrittr)
#' doc <- read_docx() %>% body_add_break()
#' print(doc, target = "body_add_break.docx" )
body_add_break <- function( x, pos = "after"){

  str <- paste0(wml_with_ns("w:p"), "<w:pPr/>",
                "<w:r><w:br w:type=\"page\"/></w:r>",
                "</w:p>")
  body_add_xml(x = x, str = str, pos = pos)

}

#' @export
#' @title add image
#' @description add an image into an rdocx object
#' @inheritParams body_add_break
#' @param src image filename
#' @param style paragraph style
#' @param width height in inches
#' @param height height in inches
#' @examples
#' doc <- read_docx()
#'
#' img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
#' }
#'
#' print(doc, target = "body_add_img.docx" )
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
body_add_img <- function( x, src, style = NULL, width, height, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$paragraph

  style_id <- x$doc_obj$get_style_id(style=style, type = "paragraph")

  ext_img <- external_img(src, width = width, height = height)
  xml_elt <- format(ext_img, type = "wml")
  xml_elt <- paste0(wml_with_ns("w:p"),
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
                    xml_elt,
                    "</w:p>")

  x <- docx_reference_img(x, src)
  xml_elt <- wml_link_images( x, xml_elt )


  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add ggplot
#' @description add a ggplot as a png image into an rdocx object
#' @inheritParams body_add_break
#' @param value ggplot object
#' @param style paragraph style
#' @param width height in inches
#' @param height height in inches
#' @param ... Arguments to be passed to png function.
#' @importFrom grDevices png dev.off
#' @import ggplot2
#' @examples
#' library(ggplot2)
#'
#' doc <- read_docx()
#'
#' gg_plot <- ggplot(data = iris ) +
#'   geom_point(mapping = aes(Sepal.Length, Petal.Length))
#'
#' if( capabilities(what = "png") )
#'   doc <- body_add_gg(doc, value = gg_plot, style = "centered" )
#'
#' print(doc, target = "body_add_gg.docx" )
body_add_gg <- function( x, value, width = 6, height = 5, style = NULL, ... ){
  stopifnot(inherits(value, "gg") )
  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')
  png(filename = file, width = width, height = height, units = "in", res = 300, ...)
  print(value)
  dev.off()
  on.exit(unlink(file))
  body_add_img(x, src = file, style = style, width = width, height = height)
}

#' @export
#' @title add paragraph of text
#' @description add a paragraph of text into an rdocx object
#' @param x a docx device
#' @param value a character
#' @param style paragraph style
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_par("A title", style = "heading 1") %>%
#'   body_add_par("Hello world!", style = "Normal") %>%
#'   body_add_par("centered text", style = "centered")
#'
#' print(doc, target = "body_add_par.docx" )
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
#' @importFrom htmltools htmlEscape
body_add_par <- function( x, value, style = NULL, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$paragraph

  style_id <- x$doc_obj$get_style_id(style=style, type = "paragraph")

  xml_elt <- paste0(wml_with_ns("w:p"),
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr><w:r><w:t xml:space=\"preserve\">",
                    htmlEscape(value), "</w:t></w:r></w:p>")
  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add fpar
#' @description add an \code{fpar} (a formatted paragraph) into an rdocx object
#' @param x a docx device
#' @param value a character
#' @param style paragraph style
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' library(magrittr)
#' bold_face <- shortcuts$fp_bold(font.size = 30)
#' bold_redface <- update(bold_face, color = "red")
#' fpar_ <- fpar(ftext("Hello ", prop = bold_face),
#'               ftext("World", prop = bold_redface ),
#'               ftext(", how are you?", prop = bold_face ) )
#' doc <- read_docx() %>% body_add_fpar(fpar_)
#'
#' print(doc, target = "body_add_fpar.docx" )
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
#' @seealso \code{\link{fpar}}
body_add_fpar <- function( x, value, style = NULL, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$paragraph
  style_id <- x$doc_obj$get_style_id(style=style, type = "paragraph")

  xml_elt <- format(value, type = "wml")
  xml_elt <- gsub("<w:p>", wml_with_ns("w:p"), xml_elt )
  xml_node <- as_xml_document(xml_elt)
  ppr <- xml_child(xml_node, "w:pPr")
  xml_remove(xml_children(ppr))
  xml_add_child(ppr,
                as_xml_document(
                  paste0(
                    "<w:pStyle xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:val=\"",
                    style_id, "\"/>"))
                )


  body_add_xml(x = x, str = as.character(xml_node), pos = pos)
}

#' @export
#' @title add table
#' @description add a table into an rdocx object
#' @param x a docx device
#' @param value a data.frame
#' @param style table style
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param header display header if TRUE
#' @param first_row,last_row,first_column,last_column,no_hband,no_vband logical for Word table options
#' @examples
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_table(iris, style = "table_template")
#'
#' print(doc, target = "body_add_table.docx" )
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
body_add_table <- function( x, value, style = NULL, pos = "after", header = TRUE,
                            first_row = TRUE, first_column = FALSE,
                            last_row = FALSE, last_column = FALSE,
                            no_hband = FALSE, no_vband = TRUE ){

  stopifnot(is.data.frame(value))

  if( is.null(style) )
    style <- x$default_styles$table

  style_id <- x$doc_obj$get_style_id(style=style, type = "table")

  value <- characterise_df(value)

  xml_elt <- wml_table(value, style_id,
            first_row, last_row,
            first_column, last_column,
            no_hband, no_vband, header)

  body_add_xml(x = x, str = xml_elt, pos = pos)
}



#' @export
#' @title add table of content
#' @description add a table of content into an rdocx object
#' @param x an rdocx object
#' @param level max title level of the table
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param style optional. style in the document that will be used to build entries of the TOC.
#' @param separator optional. Some configurations need "," (i.e. from Canada) separator instead of ";"
#' @examples
#' library(magrittr)
#' doc <- read_docx() %>% body_add_toc()
#'
#' print(doc, target = "body_add_toc.docx" )
body_add_toc <- function( x, level = 3, pos = "after", style = NULL, separator = ";"){

  if( is.null( style )){
    str <- paste0(wml_with_ns("w:p"), "<w:pPr/>",
                  "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>",
                  "<w:r><w:instrText xml:space=\"preserve\" w:dirty=\"true\">TOC \u005Co &quot;1-%.0f&quot; \u005Ch \u005Cz \u005Cu</w:instrText></w:r>",
                  "<w:r><w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/></w:r>",
                  "</w:p>")
    out <- sprintf(str, level)
  } else {
    str <- paste0(wml_with_ns("w:p"), "<w:pPr/>",
                  "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>",
                  "<w:r><w:instrText xml:space=\"preserve\" w:dirty=\"true\">TOC \u005Ch \u005Cz \u005Ct \"%s%s1\"</w:instrText></w:r>",
                  "<w:r><w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/></w:r>",
                  "</w:p>")
    out <- sprintf(str, style, separator)
  }

  body_add_xml(x = x, str = out, pos = pos)

}



#' @export
#' @title add an xml string as document element
#' @description Add an xml string as document element in the document. This function
#' is to be used to add custom openxml code.
#' @param x an rdocx object
#' @param str a wml string
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @importFrom xml2 xml_replace xml_add_sibling
body_add_xml <- function(x, str, pos){
  xml_elt <- as_xml_document(str)
  pos_list <- c("after", "before", "on")

  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  cursor_elt <- x$doc_obj$get_at_cursor()

  if( length(x) == 1 ){
    xml_add_sibling(cursor_elt, xml_elt, .where = "before")
    x <- cursor_end(x)
  } else {
    if( pos != "on")
      xml_add_sibling(cursor_elt, xml_elt, .where = pos)
    else {
      xml_replace(cursor_elt, xml_elt)
    }
    if(pos == "after")
      x <- cursor_forward(x)
  }


  x
}




#' @export
#' @importFrom uuid UUIDgenerate
#' @title add bookmark
#' @description Add a bookmark at the cursor location.
#' @param x an rdocx object
#' @param id bookmark name
#' @examples
#'
#' # cursor_bookmark ----
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_par("centered text", style = "centered") %>%
#'   body_bookmark("text_to_replace")
body_bookmark <- function(x, id){
  cursor_elt <- x$doc_obj$get_at_cursor()
  ns_ <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
  new_id <- UUIDgenerate()

  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\" %s/>", new_id, id, ns_ )
  bm_start_end <- sprintf("<w:bookmarkEnd %s w:id=\"%s\"/>", ns_, new_id )

  path_ <- paste0( xml_path(cursor_elt), "//w:r")

  node <- xml_find_first(x$doc_obj$get(), path_ )
  xml_add_sibling(node, as_xml_document( bm_start_str ) , .where = "before")
  xml_add_sibling(node, as_xml_document( bm_start_end ) , .where = "after")

  x
}



#' @export
#' @title remove an element
#' @description remove element pointed by cursor from a Word document
#' @importFrom xml2 xml_remove
#' @param x an rdocx object
#' @examples
#' library(officer)
#' library(magrittr)
#'
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>%
#'   rep(20) %>% paste(collapse = "")
#' str2 <- "Drop that text"
#' str3 <- "Aenean venenatis varius elit et fermentum vivamus vehicula. " %>%
#'   rep(20) %>% paste(collapse = "")
#'
#' my_doc <- read_docx()  %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   body_add_par(value = str2, style = "centered") %>%
#'   body_add_par(value = str3, style = "Normal")
#'
#' print(my_doc, target = "init_doc.docx")
#'
#' my_doc <- read_docx(path = "init_doc.docx")  %>%
#'   cursor_reach(keyword = "that text") %>%
#'   body_remove()
#'
#' print(my_doc, target = "result_doc.docx")
body_remove <- function(x){

  cursor_elt <- x$doc_obj$get_at_cursor()

  if( x$doc_obj$length() == 1 && xml_name(cursor_elt) == "sectPr"){
    warning("There is nothing left to remove in the document")
    return(x)
  }

  x$doc_obj$cursor_forward()
  new_cursor_elt <- x$doc_obj$get_at_cursor()
  xml_remove(cursor_elt)
  x$doc_obj$set_cursor(xml_path(new_cursor_elt))
  x
}

#' @export
#' @importFrom purrr is_scalar_character
#' @title replace text at a bookmark location
#' @description replace text content enclosed in a bookmark
#' by another text. A bookmark will be considered as valid if enclosing words
#' within a paragraph, i.e. a bookmark along two or more paragraphs is invalid,
#' a bookmark set on a whole paragraph is also invalid, bookmarking few words inside a paragraph
#' is valid.
#' @param x a docx device
#' @param bookmark bookmark id
#' @param value a character
#' @examples
#' library(magrittr)
#' doc <- read_docx() %>%
#'   body_add_par("centered text", style = "centered") %>%
#'   slip_in_text(". How are you", style = "strong") %>%
#'   body_bookmark("text_to_replace") %>%
#'   body_replace_at("text_to_replace", "not left aligned")
body_replace_at <- function( x, bookmark, value ){
  stopifnot(is_scalar_character(value), is_scalar_character(bookmark))
  x$doc_obj$cursor_replace_first_text(bookmark, value)
  x
}


#' @export
#' @title add section
#' @description add a section in a Word document. A section has effect
#' on preceding paragraphs or tables.
#'
#' @details
#' A section start at the end of the previous section (or the beginning of
#' the document if no preceding section exists), it stops where the section is declared.
#' The function \code{body_end_section()} is reflecting that Word concept.
#' The function \code{body_default_section()} is only modifying the default section of
#' the document.
#' @importFrom xml2 xml_remove
#' @param x an rdocx object
#' @param landscape landscape orientation
#' @param colwidths columns widths in percent, if 3 values, 3 columns will be produced.
#' Sum of this argument should be 1.
#' @param space space in percent between columns.
#' @param sep if TRUE a line is sperating columns.
#' @param continuous TRUE for a continuous section break.
#' @examples
#' library(magrittr)
#'
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>%
#'   rep(10) %>% paste(collapse = "")
#'
#' my_doc <- read_docx() %>%
#'   # add a paragraph
#'   body_add_par(value = str1, style = "Normal") %>%
#'   # add a continuous section
#'   body_end_section(continuous = TRUE) %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   # preceding paragraph is on a new column
#'   break_column_before() %>%
#'   # add a two columns continous section
#'   body_end_section(colwidths = c(.6, .4),
#'                    space = .05, sep = FALSE, continuous = TRUE) %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   # add a continuous section ... so far there is no break page
#'   body_end_section(continuous = TRUE) %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   body_default_section(landscape = TRUE)
#'
#' print(my_doc, target = "section.docx")
body_end_section <- function(x, landscape = FALSE, colwidths = c(1), space = .05, sep = FALSE, continuous = FALSE){

  stopifnot(all.equal( sum(colwidths), 1 ) )

  if( landscape && continuous ){
    stop("using landscape=TRUE and continuous=TRUE is not possible as changing orientation require a new page.")
  }

  sdim <- x$sect_dim
  h_ref <- sdim$page["height"];w_ref <- sdim$page["width"]
  mar_t <- sdim$margins["top"];mar_b <- sdim$margins["bottom"]
  mar_r <- sdim$margins["right"];mar_l <- sdim$margins["left"]
  mar_h <- sdim$margins["header"];mar_f <- sdim$margins["footer"]

  if( !landscape ){
    h <- h_ref
    w <- w_ref
    mar_top <- mar_t
    mar_bottom <- mar_b
    mar_right <- mar_r
    mar_left <- mar_l
  } else {
    h <- w_ref
    w <- h_ref
    mar_top <- mar_r
    mar_bottom <- mar_l
    mar_right <- mar_t
    mar_left <- mar_b
  }
  pgsz_str <- "<w:pgSz %sw:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, ifelse( landscape, "w:orient=\"landscape\" ", ""), w, h )

  mar_str <- "<w:pgMar w:top=\"%.0f\" w:right=\"%.0f\" w:bottom=\"%.0f\" w:left=\"%.0f\" w:header=\"%.0f\" w:footer=\"%.0f\" w:gutter=\"0\"/>"
  mar_str <- sprintf(mar_str, mar_top, mar_right, mar_bottom, mar_left, mar_h, mar_f )

  width_ <- w - mar_right - mar_left
  column_values <- colwidths - space
  columns_str_all_but_last <- sprintf("<w:col w:w=\"%.0f\" w:space=\"%.0f\"/>",
                                      column_values[-length(column_values)] * width_,
                                      space * width_)
  columns_str_last <- sprintf("<w:col w:w=\"%.0f\"/>",
                              column_values[length(column_values)] * width_)
  columns_str <- c(columns_str_all_but_last, columns_str_last)
  if( length(colwidths) > 1 )
    columns_str <- sprintf("<w:cols w:num=\"%.0f\" w:sep=\"%.0f\" w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
          length(colwidths), as.integer(sep), space * w, paste0(columns_str, collapse = "") )
  else columns_str <- sprintf("<w:cols w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
                              space * w, paste0(columns_str, collapse = "") )

  str <- paste0( wml_with_ns("w:p"),
                 "<w:pPr><w:sectPr>",
                 ifelse( continuous, "<w:type w:val=\"continuous\"/>", "" ),
                 pgsz_str, mar_str, columns_str, "</w:sectPr></w:pPr></w:p>")
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @rdname body_end_section
body_default_section <- function(x, landscape = FALSE){

  sdim <- x$sect_dim
  h_ref <- sdim$page["height"];w_ref <- sdim$page["width"]
  mar_t <- sdim$margins["top"];mar_b <- sdim$margins["bottom"]
  mar_r <- sdim$margins["right"];mar_l <- sdim$margins["left"]
  mar_h <- sdim$margins["header"];mar_f <- sdim$margins["footer"]

  if( !landscape ){
    h <- h_ref
    w <- w_ref
    mar_top <- mar_t
    mar_bottom <- mar_b
    mar_right <- mar_r
    mar_left <- mar_l
  } else {
    h <- w_ref
    w <- h_ref
    mar_top <- mar_r
    mar_bottom <- mar_l
    mar_right <- mar_t
    mar_left <- mar_b
  }
  pgsz_str <- "<w:pgSz %sw:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, ifelse( landscape, "w:orient=\"landscape\" ", ""), w, h )

  mar_str <- "<w:pgMar w:top=\"%.0f\" w:right=\"%.0f\" w:bottom=\"%.0f\" w:left=\"%.0f\" w:header=\"%.0f\" w:footer=\"%.0f\" w:gutter=\"0\"/>"
  mar_str <- sprintf(mar_str, mar_top, mar_right, mar_bottom, mar_left, mar_h, mar_f )

  str <- paste0( wml_with_ns("w:sectPr"),
                 "<w:type w:val=\"continuous\"/>",
                 pgsz_str, mar_str, "</w:sectPr>")
  last_sect <- xml_find_first(x$doc_obj$get(), "/w:document/w:body/w:sectPr[last()]")
  xml_replace(last_sect, as_xml_document(str) )
  x
}

#' @export
#' @rdname body_end_section
break_column_before <- function( x ){
  xml_elt <- paste0( wml_with_ns("w:r"), "<w:br w:type=\"column\"/>", "</w:r>")
  slip_in_xml(x = x, str = xml_elt, pos = "before")
}


