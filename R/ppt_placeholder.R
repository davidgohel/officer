simple_shape <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">
  <p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>"

empty_shape <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">
  <p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr/></p:sp>"

r_shape <- paste0( "<a:r xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                   "%s",# style formatted
                   "<a:t>%s</a:t>",#text
                   "</a:r>" )

pml_ns <- " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""

get_shape_id <- function(x, type = NULL, id_chr = NULL ){
  shape_index_data <- slide_summary(x)
  shape_index_data$shape_id <- seq_len(nrow(shape_index_data))

  if( !is.null(type) && !is.null(id_chr) ){
    filter_criteria <- interp(~ type == tp & id == index, tp = type, index = id_chr)
    shape_index_data <- filter_(shape_index_data, filter_criteria)
  } else if( is.null(type) && !is.null(id_chr) ){
    filter_criteria <- interp(~ id == index, index = id_chr)
    shape_index_data <- filter_(shape_index_data, filter_criteria)
  } else if( !is.null(type) && is.null(id_chr) ){
    filter_criteria <- interp(~ type == tp, tp = type)
    shape_index_data <- filter_(shape_index_data, filter_criteria)
  } else {
    filter_criteria <- interp(~ type == tp, tp = type)
    shape_index_data <- shape_index_data[nrow(shape_index_data), ]
  }

  if( nrow(shape_index_data) < 1 )
    stop("selection does not match any row in slide_summary. Use function slide_summary.", call. = FALSE)
  else if( nrow(shape_index_data) > 1 )
    stop("selection does match more than a single row in slide_summary. Use function slide_summary.", call. = FALSE)

  shape_index_data$shape_id
}

#' @export
#' @title remove shape
#' @description remove a shape in a slide
#' @param x a pptx device
#' @param type placeholder type
#' @param id_chr placeholder id (a string). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body'.
#' Values can be read from \code{\link{slide_summary}}.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- placeholder_set_text(x = doc, type = "title", str = "Un titre")
#' slide_summary(doc) # read column id here
#' doc <- placeholder_remove(x = doc, type = "title", id_chr = "2")
#' print(doc, target = fileout )
#' @importFrom xml2 xml_remove xml_find_all
placeholder_remove <- function( x, type = NULL, id_chr = NULL ){

  slide <- x$slide$get_slide(x$cursor)
  shape_id <- get_shape_id(x, type = type, id_chr = id_chr )
  str = "p:cSld/p:spTree/*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]"
  xml_remove(xml_find_all(slide$get(), str)[[shape_id]])

  slide$save()
  x$slide$update()
  x
}


#' @export
#' @title add text into a new shape
#' @description add text into a new shape in a slide
#' @param x a pptx device
#' @param str text to add
#' @param type placeholder type
#' @param index placeholder index (integer). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body'.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- placeholder_set_text(x = doc, type = "title", str = "Un titre")
#' doc <- placeholder_set_text(x = doc, type = "ftr", str = "pied de page")
#' doc <- placeholder_set_text(x = doc, type = "dt", str = format(Sys.Date()))
#' doc <- placeholder_set_text(x = doc, type = "sldNum", str = "slide 1")#
#'
#' doc <- add_slide(doc, layout = "Title Slide", master = "Office Theme")
#' doc <- placeholder_set_text(x = doc, type = "subTitle", str = "Un sous titre")
#' doc <- placeholder_set_text(x = doc, type = "ctrTitle", str = "Un titre")
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
placeholder_set_text <- function( x, str, type = "title", index = 1 ){

  stopifnot( type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = type, index = index)
  xml_elt <- sprintf( simple_shape, xfrm_df$ph, str )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))

  slide$save()
  x$slide$update()
  x
}

#' @export
#' @title add a new shape
#' @description add a new shape in a slide
#' @param x a pptx device
#' @param type placeholder type
#' @param index placeholder index (integer). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body'.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- placeholder_set(x = doc, type = "title")
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
placeholder_set <- function( x, type = "title", index = 1 ){

  stopifnot( type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = type, index = index)
  xml_elt <- sprintf( empty_shape, xfrm_df$ph )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))

  slide$save()
  x$slide$update()
  x
}


#' @export
#' @title append text
#' @description append text into a paragraph of a pptx object
#' @inheritParams placeholder_remove
#' @param str text to add
#' @param style text style, a \code{\link{fp_text}} object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @examples
#' library(magrittr)
#' fileout <- tempfile(fileext = ".pptx")
#' my_pres <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   placeholder_set(type = "body")
#'
#' small_red <- fp_text(color = "red", font.size = 14)
#'
#' my_pres <- my_pres %>%
#'   placeholder_add_paragraph(level = 3) %>%
#'   placeholder_add_text(str = "A small red text.", style = small_red) %>%
#'   placeholder_add_paragraph(level = 2) %>%
#'   placeholder_add_text(str = "Level 2")
#'
#' print(my_pres, target = fileout)
#' @importFrom xml2 xml_child xml_children xml_add_child
placeholder_add_text <- function( x, str, type = NULL, id_chr = NULL,
  style = fp_text(font.size = 0), pos = "after" ){

  slide <- x$slide$get_slide(x$cursor)
  shape_id <- get_shape_id(x, type = type, id_chr = id_chr )

  nodes <- xml_find_all(slide$get(), "p:cSld/p:spTree/p:sp")

  current_p <- xml_child(nodes[[shape_id]], "/a:p[last()]")
  if( inherits(current_p, "xml_missing") )
    stop("Could not find any paragraph in the selected shape.")
  r_shape_ <- sprintf(r_shape, format(style, type = "pml"), str )

  if( pos == "after" )
    where_ <- length(xml_children(current_p))
  else where_ <- 0

  xml_add_child(current_p, as_xml_document(r_shape_), .where = where_ )

  slide$save()
  x$slide$update()

  x
}

#' @export
#' @title append text
#' @description append text into a paragraph of a pptx object
#' @inheritParams placeholder_remove
#' @param level paragraph level
#' @examples
#' library(magrittr)
#'
#' fileout <- tempfile(fileext = ".pptx")
#' default_text <- fp_text(font.size = 0, bold = TRUE, color = "red")
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   placeholder_set_text(type = "body", str = "A text") %>%
#'   placeholder_add_paragraph(level = 2) %>%
#'   placeholder_add_text(str = "and another, ", style = default_text ) %>%
#'   placeholder_add_paragraph(level = 3) %>%
#'   placeholder_add_text(str = "and another!",
#'     style = update(default_text, color = "blue")) %>%
#'   print(target = fileout)
#' @importFrom xml2 xml_child xml_children xml_add_child
placeholder_add_paragraph <- function( x, type = NULL, id_chr = NULL, level = 1 ){

  slide <- x$slide$get_slide(x$cursor)
  shape_id <- get_shape_id(x, type = type, id_chr = id_chr )

  nodes <- xml_find_all(slide$get(), "p:cSld/p:spTree/p:sp")

  current_p <- xml_child(nodes[[shape_id]], "/p:txBody")

  if( inherits(current_p, "xml_missing") ){
    if( level > 1 )
      p_shape <- sprintf("<a:p><a:pPr lvl=\"%.0f\"/></a:p>", level - 1)
    else
      p_shape <- "<a:p/>"
    simple_shape <- sprintf( paste0( "<p:txBody%s><a:bodyPr/><a:lstStyle/>",
                            p_shape, "</p:txBody>"), pml_ns)
    xml_add_child(nodes[[shape_id]], as_xml_document(simple_shape) )
  } else {
    if( level > 1 )
      simple_shape <- sprintf("<a:p%s><a:pPr lvl=\"%.0f\"/></a:p>", pml_ns, level - 1)
    else
      simple_shape <- sprintf("<a:p%s/>", pml_ns)
    xml_add_child(current_p, as_xml_document(simple_shape) )
  }


  slide$save()
  x$slide$update()

  x
}


#' @export
#' @title add a string as xml
#' @description add a string (valid xml) into a pptx object
#' @param x a pptx device
#' @param type placeholder type
#' @param index placeholder index (integer). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body'.
#' @param value a character
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
placeholder_set_xml <- function( x, value, type = "body", index = 1 ){

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


#' @export
#' @title add image into a new shape
#' @description add an image into a new shape in a slide
#' @param x a pptx device
#' @param type placeholder type
#' @param index placeholder index (integer). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body'.
#' @param src image path
#' @param width,height image size in inches
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'
#' img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- placeholder_set_img(x = doc, type = "body", src = img.file, height = 1.06, width = 1.39 )
#' }
#' if( require("ionicons") ){
#'   calendar_src = as_png(name = "calendar", fill = "#FFE64D", width = 144, height = 144)
#'   doc <- placeholder_set_img(x = doc, type = "dt", src = calendar_src )
#' }
#' if( require("devEMF") ){
#'   emf("bar.emf")
#'   barplot(1:10, col = 1:10)
#'   dev.off()
#'   doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'   doc <- placeholder_set_img(x = doc, type = "body", src = "bar.emf" )
#' }
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
placeholder_set_img <- function( x, src, type = "body", index = 1, width = NULL, height = NULL ){

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

