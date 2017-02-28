#' @export
#' @title add a new empty shape
#' @description add a new empty shape in the current slide.
#' @param x a pptx device
#' @param type placeholder type
#' @param index placeholder index (integer). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body'.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_empty(x = doc, type = "title")
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
ph_empty <- function( x, type = "title", index = 1 ){

  stopifnot( type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = type, index = index)

  empty_shape <- paste0(pml_with_ns("p:sp"),
                        "<p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr/></p:sp>")

  xml_elt <- sprintf( empty_shape, xfrm_df$ph )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))

  slide$save()
  x$slide$update()
  x
}

#' @export
#' @title add text into a new shape
#' @description add text into a new shape in a slide.
#' @inheritParams ph_empty
#' @param str text to add
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with_text(x = doc, type = "title", str = "Un titre")
#' doc <- ph_with_text(x = doc, type = "ftr", str = "pied de page")
#' doc <- ph_with_text(x = doc, type = "dt", str = format(Sys.Date()))
#' doc <- ph_with_text(x = doc, type = "sldNum", str = "slide 1")
#'
#' doc <- add_slide(doc, layout = "Title Slide", master = "Office Theme")
#' doc <- ph_with_text(x = doc, type = "subTitle", str = "Un sous titre")
#' doc <- ph_with_text(x = doc, type = "ctrTitle", str = "Un titre")
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
ph_with_text <- function( x, str, type = "title", index = 1 ){

  stopifnot( type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = type, index = index)
  xml_elt <- pml_shape_str(str = str, ph = xfrm_df$ph )
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))

  slide$save()
  x$slide$update()
  x
}


#' @export
#' @title add table
#' @description add a table as a new shape in the current slide.
#' @inheritParams ph_empty
#' @param value data.frame
#' @param first_row,last_row,first_column,last_column logical for PowerPoint table options
#' @examples
#' library(magrittr)
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with_table(value = mtcars[1:6,], type = "body",
#'     last_row = FALSE, last_column = FALSE, first_row = TRUE)
#'
#' print(doc, target = "ph_with_table.pptx")
ph_with_table <- function( x, value, type = "title", index = 1,
                           first_row = TRUE, first_column = FALSE,
                           last_row = FALSE, last_column = FALSE ){

  stopifnot( type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = type, index = index)

  def_height <- xfrm_df$cy*914400 / (nrow(value) + 1)
  def_width <- xfrm_df$cx*914400 / (ncol(value))


  style_id <- x$table_styles$def[1]
  xml_elt <- pml_table(value, style_id = style_id,
                       col_width = as.integer(def_width),
                       row_height = as.integer(def_height),
                       first_row = first_row, last_row = last_row,
                       first_column = first_column, last_column = last_column )

  xml_elt <- paste0(
    "<p:graphicFrame ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
    "<p:nvGraphicFramePr>",
    "<p:cNvPr id=\"\" name=\"\"/>",
    "<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"true\"/></p:cNvGraphicFramePr>",
    "<p:nvPr/>",
    "</p:nvGraphicFramePr>",
    "<p:xfrm rot=\"0\">",
    sprintf( "<a:off x=\"%.0f\" y=\"%.0f\"/>", xfrm_df$offx*914400, xfrm_df$offy*914400 ),
    sprintf( "<a:ext cx=\"%.0f\" cy=\"%.0f\"/>", xfrm_df$cx*914400, xfrm_df$cy*914400 ),
    "</p:xfrm>",
    "<a:graphic>",
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">",
    xml_elt,
    "</a:graphicData>",
    "</a:graphic>",
    "</p:graphicFrame>"
  )
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))
  slide$save()
  x$slide$update()
  x
}



#' @export
#' @title add image
#' @description add an image as a new shape in the current slide.
#' @inheritParams ph_empty
#' @param src image path
#' @param width,height image size in inches
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'
#' img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- ph_with_img(x = doc, type = "body", src = img.file, height = 1.06, width = 1.39 )
#' }
#' if( require("ionicons") ){
#'   calendar_src = as_png(name = "calendar", fill = "#FFE64D", width = 144, height = 144)
#'   doc <- ph_with_img(x = doc, type = "dt", src = calendar_src )
#' }
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
ph_with_img <- function( x, src, type = "body", index = 1, width = NULL, height = NULL ){

  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm(type = type, index = index)

  if( is.null(width)) width <- xfrm$cx
  if( is.null(height)) height <- xfrm$cy

  ext_img <- external_img(src, width = width, height = height)
  xml_elt <- format(ext_img, type = "pml")

  slide$reference_img(src = src, dir_name = file.path(x$package_dir, "ppt/media"))
  xml_elt <- fortify_pml_images(x, xml_elt)

  doc <- as_xml_document(xml_elt)

  node <- xml_find_first( doc, "p:spPr")
  off <- xml_child(node, "a:xfrm/a:off")
  xml_attr( off, "x") <- sprintf( "%.0f", xfrm$offx * 914400 )
  xml_attr( off, "y") <- sprintf( "%.0f", xfrm$offy * 914400 )

  xmlslide <- slide$get()


  xml_add_child(xml_find_first(xmlslide, "p:cSld/p:spTree"), doc)
  slide$save()
  x$slide$update()
  x

}

fortify_pml_images <- function(x, str){

  slide <- x$slide$get_slide(x$cursor)
  ref <- slide$rel_df()

  ref <- filter_(ref, interp(~ ext_src != "") )
  doc <- as_xml_document(str)
  for(id in seq_along(ref$ext_src) ){
    xpth <- paste0("//p:pic/p:blipFill/a:blip",
                   sprintf( "[contains(@r:embed,'%s')]", ref$ext_src[id]),
                   "")

    src_nodes <- xml_find_all(doc, xpth)
    xml_attr(src_nodes, "r:embed") <- ref$id[id]
  }
  as.character(doc)
}


#' @export
#' @title add an xml string as new shape
#' @description Add an xml string as new shape in the current slide. This function
#' is to be used to add custom openxml code.
#' @inheritParams ph_empty
#' @param value a character
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
ph_from_xml <- function( x, value, type = "body", index = 1 ){

  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm(type = type, index = index)

  doc <- as_xml_document(value)
  node <- xml_find_first( doc, "//*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]")
  node <- set_xfrm_attr(node, offx = xfrm$offx, offy = xfrm$offy,
                        cx = xfrm$cx, cy = xfrm$cy)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), doc)

  slide$fortify_id()$save()
  x$slide$update()
  x
}



