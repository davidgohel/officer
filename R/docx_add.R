#' @export
#' @title add an image
#' @description add an image into a docx object
#' @param x a docx device
#' @param src image filename
#' @param style paragraph style
#' @param width height in inches
#' @param height height in inches
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' doc <- docx()
#'
#' img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- docx_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
#' }
#' if( require("ionicons") ){
#'   calendar_src = as_png(name = "calendar", fill = "#FFE64D", width = 144, height = 144)
#'   doc <- docx_add_img(x = doc, src = calendar_src, height = 2, width = 2 )
#' }
#' if( require("devEMF") ){
#'   emf("bar.emf", height = 5, width = 5)
#'   barplot(1:10, col = 1:10)
#'   dev.off()
#'   doc <- docx_add_img(x = doc, src = "bar.emf", height = 5, width = 5)
#' }
#'
#' print(doc, target = "docx_add_img.docx" )
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
docx_add_img <- function( x, src, style = "Normal", width, height, pos = "after" ){

  style_id <- get_style_id(x=x, style=style, type = "paragraph")

  ext_img <- external_img(src, width = width, height = height)
  xml_elt <- format(ext_img, type = "wml")
  xml_elt <- paste0("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
                    xml_elt,
                    "</w:p>")

  x <- docx_reference_img(x, src)
  xml_elt <- wml_link_images( x, xml_elt )


  add_xml_node(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add a paragraph
#' @description add a paragraph into a docx object
#' @param x a docx device
#' @param value a character
#' @param style paragraph style
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
docx_add_par <- function( x, value, style, pos = "after" ){

  style_id <- get_style_id(x=x, style=style, type = "paragraph")

  xml_elt <- paste0("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr><w:r><w:t xml:space=\"preserve\">",
                    value, "</w:t></w:r></w:p>")
  add_xml_node(x = x, str = xml_elt, pos = pos)
}

as_tc <- function(x, collapse = FALSE ){
  str <- paste0("<w:tc><w:trPr/><w:p><w:r><w:t>", gsub("(^[ ]|[ ]$)", "", format(x)), "</w:t></w:r></w:p></w:tc>")
  if( collapse )
    str <- paste(str, collapse = "")
  str
}

#' @export
#' @title add a table
#' @description add a table into a docx object
#' @param x a docx device
#' @param value a data.frame
#' @param style table style
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param width table width for column width calculation
#' @param first_row,last_row,first_column,last_column,no_hband,no_vband logical for Word table options
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
docx_add_table <- function( x, value, style, pos = "after", width = 5,
                            first_row = TRUE, first_column = FALSE,
                            last_row = TRUE, last_column = FALSE,
                            no_hband = FALSE, no_vband = TRUE ){

  style_id <- get_style_id(x=x, style=style, type = "table")

  tbl_look <- "<w:tblLook w:firstRow=\"%.0f\" w:lastRow=\"%.0f\" w:firstColumn=\"%.0f\" w:lastColumn=\"%.0f\" w:noHBand=\"%.0f\" w:noVBand=\"%.0f\" />"
  tbl_look <- sprintf(tbl_look, first_row, last_row, first_column, last_column, no_hband, no_vband)

  dat <- lapply(value, as_tc)
  dat <- do.call(cbind, dat)
  dat <- apply( dat, 1, function(x){
    paste0("<w:tr>", paste(x, collapse = ""), "</w:tr>")
  })
  dat <- paste(dat, collapse = "")
  dat <- paste( paste0("<w:tr><w:trPr><w:tblHeader/></w:trPr>",
                       as_tc(names(value), collapse = TRUE), "</w:tr>"),
                dat )

  width <- width * 72 * 20
  grid_col <- sprintf("<w:gridCol w:w=\"%.0f\"/>", width / ncol(value) )
  grid_col <- rep(grid_col, ncol(value))
  grid_col <- paste(grid_col, collapse = "")
  grid_col <- paste0("<w:tblGrid>", grid_col, "</w:tblGrid>")

  tbpr <- "<w:tblPr><w:tblStyle w:val=\"%s\"/><w:tblW/>%s</w:tblPr>"
  tbpr <- sprintf(tbpr, style_id, tbl_look)

  xml_elt <- paste0( sprintf("<w:tbl %s>", base_ns),
          tbpr, grid_col, dat, "</w:tbl>")

  add_xml_node(x = x, str = xml_elt, pos = pos)
}



#' @export
#' @title add a table of content
#' @description add a table of content into a docx object
#' @param x a docx object
#' @param level max title level of the table
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param style optional. style in the document that will be used to build entries of the TOC.
#' @param separator optional. Some configurations need "," (i.e. from Canada) separator instead of ";"
docx_add_toc <- function( x, level = 3, pos = "after", style = NULL, separator = ";"){

  if( is.null( style )){
    str <- paste0("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr/>",
                  "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>",
                  "<w:r><w:instrText xml:space=\"preserve\" w:dirty=\"true\">TOC \u005Co &quot;1-%.0f&quot; \u005Ch \u005Cz \u005Cu</w:instrText></w:r>",
                  "<w:r><w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/></w:r>",
                  "</w:p>")
    out <- sprintf(str, level)
  } else {
    str <- paste0("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr/>",
                  "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>",
                  "<w:r><w:instrText xml:space=\"preserve\" w:dirty=\"true\">TOC \u005Ch \u005Cz \u005Ct \"%s%s1\"</w:instrText></w:r>",
                  "<w:r><w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/></w:r>",
                  "</w:p>")
    out <- sprintf(str, style, separator)
  }

  add_xml_node(x = x, str = out, pos = pos)

}



#' @export
#' @title add a page break
#' @description add a page break into a docx object
#' @param x a docx object
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
docx_add_pbreak <- function( x, pos = "after"){

  str <- paste0("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr/>",
                "<w:r><w:br w:type=\"page\"/></w:r>",
                "</w:p>")
  add_xml_node(x = x, str = str, pos = pos)

}

#' @export
#' @title add a wml string into a Word document
#' @importFrom xml2 as_xml_document xml_replace
#' @description The function add a wml string into
#' the document after, before or on a cursor location.
#' @param x a docx object
#' @param str a wml string
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
add_xml_node <- function(x, str, pos){
  xml_elt <- as_xml_document(str)
  pos_list <- c("after", "before", "on")

  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  cursor_elt <- xml_find_first(x$xml_doc, x$cursor)
  if( pos != "on")
    xml_add_sibling(cursor_elt, xml_elt, .where = pos)
  else {
    xml_replace(cursor_elt, xml_elt)
  }

  if(pos == "after")
    x <- cursor_forward(x)
  x
}
