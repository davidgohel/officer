simple_shape <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">
  <p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>"


#' @export
#' @title remove shape
#' @description remove a shape in a slide
#' @param x a pptx device
#' @param id placeholder id
#' @param index placeholder index. This is to be used when a placeholder id
#' is not unique in the current slide, e.g. two placeholders with id 'body'.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- placeholder_set_text(x = doc, id = "title", str = "Un titre")
#' doc <- placeholder_remove(x = doc, id = "title")
#' print(doc, target = fileout )
#' @importFrom xml2 xml_remove xml_find_all
placeholder_remove <- function( x, id = "title", index = 1 ){

  stopifnot( id %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)

  str = "p:cSld/p:spTree/*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic][p:nvSpPr/p:nvPr/p:ph[@type='%s']]"
  str <- sprintf(str, id)
  xml_remove(xml_find_all(slide$get(), str)[[index]])

  slide$save()
  x$slide$update()
  x
}


#' @export
#' @title add text into a new shape
#' @description add text into a new shape in a slide
#' @inheritParams placeholder_remove
#' @param str text to add
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- placeholder_set_text(x = doc, id = "title", str = "Un titre")
#' doc <- placeholder_set_text(x = doc, id = "ftr", str = "pied de page")
#' doc <- placeholder_set_text(x = doc, id = "dt", str = format(Sys.Date()))
#' doc <- placeholder_set_text(x = doc, id = "sldNum", str = "slide 1")#
#'
#' doc <- add_slide(doc, layout = "Title Slide", master = "Office Theme")
#' doc <- placeholder_set_text(x = doc, id = "subTitle", str = "Un sous titre")
#' doc <- placeholder_set_text(x = doc, id = "ctrTitle", str = "Un titre")
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
placeholder_set_text <- function( x, str, id = "title", index = 1 ){

  stopifnot( id %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = id, index = index)
  xml_elt <- sprintf( simple_shape, xfrm_df$ph, str )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))

  slide$save()
  x$slide$update()
  x
}


#' @export
#' @title append text
#' @description append text into a paragraph of a pptx object
#' @param x a pptx device
#' @param str text to add
#' @param uid unique id, use column \code{id} of the
#' result of \code{\link{summary.pptx}}
#' @param style text style, a \code{\link{fp_text}} object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @examples
#' library(magrittr)
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   placeholder_set_text(id = "body", str = "Un texte")
#' id_body <- summary(doc)$id[1]

#' placeholder_add_text(doc, str = " et un autre ", uid = id_body) %>%
#'   print(target = fileout)
#' @importFrom xml2 xml_child xml_children xml_add_child
placeholder_add_text <- function( x, str, uid, style = fp_text(font.size = 0), pos = "after" ){

  slide <- x$slide$get_slide(x$cursor)
  nodes <- xml_find_all(slide$get(), "p:cSld/p:spTree/p:sp")

  shape_id <- which( !is.na( xml_child(nodes, sprintf( "/p:cNvPr[@id='%s']", uid ) ) ) )

  current_p <- xml_child(nodes[[shape_id]], "/a:p[last()]")

  simple_shape <- paste0( "<a:r xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                          format(style, type = "pml"),
                          sprintf("<a:t>%s</a:t>", str),
                          "</a:r>" )

  where_ <- ifelse( pos == "after", length(xml_children(current_p)), 0 )
  xml_add_child(current_p, as_xml_document(simple_shape), .where = where_ )

  slide$save()
  x$slide$update()

  x
}

#' @export
#' @title append text
#' @description append text into a paragraph of a pptx object
#' @param x a pptx device
#' @param uid unique id, use column \code{id} of the
#' result of \code{\link{summary.pptx}}
#' @param level paragraph level
#' @examples
#' library(magrittr)
#'
#' fileout <- tempfile(fileext = ".pptx")
#' default_text <- fp_text(font.size = 0, bold = TRUE, color = "red")
#'
#' doc <- pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   placeholder_set_text(id = "title", str = "A title") %>%
#'   placeholder_set_text(id = "body", str = "A text")
#'
#' id_body <- summary(doc)$id[2]
#'
#' placeholder_add_paragraph(x = doc, uid = id_body, level = 2) %>%
#'   placeholder_add_text(str = "and another, ", uid = id_body,
#'     style = default_text ) %>%
#'   placeholder_add_paragraph(uid = id_body, level = 3) %>%
#'   placeholder_add_text(str = "and another!", uid = id_body,
#'     style = update(default_text, color = "blue")) %>%
#'   print(target = fileout)
#' @importFrom xml2 xml_child xml_children xml_add_child
placeholder_add_paragraph <- function( x, uid, level = 1 ){

  slide <- x$slide$get_slide(x$cursor)
  nodes <- xml_find_all(slide$get(), "p:cSld/p:spTree/p:sp")

  shape_id <- which( !is.na( xml_child(nodes, sprintf( "/p:cNvPr[@id='%s']", uid ) ) ) )

  current_p <- xml_child(nodes[[shape_id]], "/p:txBody")

  if( level > 1 )
    simple_shape <- sprintf("<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr lvl=\"%.0f\"/></a:p>", level - 1)
  else
    simple_shape <- "<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"

  xml_add_child(current_p, as_xml_document(simple_shape) )
  slide$save()
  x$slide$update()

  x
}


#' @export
#' @title add a string as xml
#' @description add a string (valid xml) into a pptx object
#' @inheritParams placeholder_remove
#' @param value a character
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
placeholder_set_xml <- function( x, value, id = "body", index = 1 ){

  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm(type = id, index = index)

  doc <- as_xml_document(value)
  node <- xml_find_first( doc, "//*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]")
  node <- set_xfrm_attr(node, offx = xfrm$offx, offy = xfrm$offy,
                       cx = xfrm$cx, cy = xfrm$cy)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), doc)

  slide$fortify_id()$save()
  x$slide$update()
  x
}


#' @export
#' @title add image into a new shape
#' @description add an image into a new shape in a slide
#' @inheritParams placeholder_remove
#' @param src image path
#' @param width,height image size in inches
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'
#' img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- placeholder_set_img(x = doc, id = "body", src = img.file, height = 1.06, width = 1.39 )
#' }
#' if( require("ionicons") ){
#'   calendar_src = as_png(name = "calendar", fill = "#FFE64D", width = 144, height = 144)
#'   doc <- placeholder_set_img(x = doc, id = "dt", src = calendar_src )
#' }
#' if( require("devEMF") ){
#'   emf("bar.emf")
#'   barplot(1:10, col = 1:10)
#'   dev.off()
#'   doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'   doc <- placeholder_set_img(x = doc, id = "body", src = "bar.emf" )
#' }
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
placeholder_set_img <- function( x, src, id, index = 1, width = NULL, height = NULL ){

  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm(type = id, index = index)

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

