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

  slide$fortify_id()
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
  sh_pr_df <- slide$get_xfrm(type = type, index = index)

  sh_pr_df$str <- str
  xml_elt <- do.call(pml_shape_str, sh_pr_df)
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
}


#' @export
#' @title add table
#' @description add a table as a new shape in the current slide.
#' @inheritParams ph_empty
#' @param value data.frame
#' @param header display header if TRUE
#' @param first_row,last_row,first_column,last_column logical for PowerPoint table options
#' @examples
#' library(magrittr)
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with_table(value = mtcars[1:6,], type = "body",
#'     last_row = FALSE, last_column = FALSE, first_row = TRUE)
#'
#' print(doc, target = tempfile(fileext = ".pptx"))
ph_with_table <- function( x, value, type = "body", index = 1,
                           header = TRUE,
                           first_row = TRUE, first_column = FALSE,
                           last_row = FALSE, last_column = FALSE ){
  stopifnot(is.data.frame(value))

  stopifnot( type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = type, index = index)

  xml_elt <- table_shape(x = x, value = value, left = xfrm_df$offx, top = xfrm_df$offy, width = xfrm_df$cx, height = xfrm_df$cy,
                          first_row = first_row, first_column = first_column,
                          last_row = last_row, last_column = last_column, header = header )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))
  slide$fortify_id()
  x
}



#' @export
#' @title add image
#' @description add an image as a new shape in the current slide.
#' @inheritParams ph_empty
#' @param src image filename, the basename of the file must not contain any blank.
#' @param width,height image size in inches
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- ph_with_img(x = doc, type = "body", src = img.file, height = 1.06, width = 1.39 )
#' }
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
ph_with_img <- function( x, src, type = "body", index = 1, width = NULL, height = NULL ){

  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm(type = type, index = index)

  new_src <- tempfile( fileext = gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", src) )
  file.copy( src, to = new_src )

  if( is.null(width)) width <- xfrm$cx
  else width <- width * 914400
  if( is.null(height)) height <- xfrm$cy
  else height <- height * 914400
  ext_img <- external_img(new_src, width = width / 914400, height = height / 914400)
  xml_elt <- format(ext_img, type = "pml")

  slide$reference_img(src = new_src, dir_name = file.path(x$package_dir, "ppt/media"))
  xml_elt <- fortify_pml_images(x, xml_elt)

  doc <- as_xml_document(xml_elt)

  node <- xml_find_first( doc, "p:spPr")
  off <- xml_child(node, "a:xfrm/a:off")
  xml_attr( off, "x") <- sprintf( "%.0f", xfrm$offx )
  xml_attr( off, "y") <- sprintf( "%.0f", xfrm$offy )

  xmlslide <- slide$get()


  xml_add_child(xml_find_first(xmlslide, "p:cSld/p:spTree"), doc)
  slide$fortify_id()
  x

}

#' @export
#' @title add ggplot to a pptx presentation
#' @description add a ggplot as a png image into an rpptx object
#' @inheritParams ph_empty
#' @param value ggplot object
#' @param width,height image size in inches
#' @param ... Arguments to be passed to png function.
#' @importFrom grDevices png dev.off
#' @examples
#' if( require("ggplot2") ){
#'   doc <- read_pptx()
#'   doc <- add_slide(doc, layout = "Title and Content",
#'     master = "Office Theme")
#'
#'   gg_plot <- ggplot(data = iris ) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length), size = 3) +
#'     theme_minimal()
#'
#'   if( capabilities(what = "png") ){
#'     doc <- ph_with_gg(doc, value = gg_plot )
#'     doc <- ph_with_gg_at(doc, value = gg_plot,
#'       height = 4, width = 8, left = 4, top = 4 )
#'   }
#'
#'   print(doc, target = tempfile(fileext = ".pptx"))
#' }
ph_with_gg <- function( x, value, type = "body", index = 1, width = NULL, height = NULL, ... ){

  if( !requireNamespace("ggplot2") )
    stop("package ggplot2 is required to use this function")

  slide <- x$slide$get_slide(x$cursor)
  xfrm_df <- slide$get_xfrm(type = type, index = index)
  if( is.null( width ) ) width <- xfrm_df$cx / 914400
  if( is.null( height ) ) height <- xfrm_df$cy / 914400

  stopifnot(inherits(value, "gg") )
  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')
  png(filename = file, width = width, height = height, units = "in", res = 300, ...)
  print(value)
  dev.off()
  on.exit(unlink(file))
  ph_with_img( x, src = file, type = type, index = index,
               width = width, height = height )
}

#' @title add unordered list to a pptx presentation
#' @description add an unordered list of text
#' into an rpptx object. Each text is associated with
#' a hierarchy level.
#'
#' @param x rpptx object
#' @param type placeholder type
#' @param index placeholder index (integer). This is to be used
#' when a placeholder type is not unique in the current slide,
#' e.g. two placeholders with type 'body'.
#' @param str_list list of strings to be included in the object
#' @param level_list list of levels for hierarchy structure
#' @param style text style, a \code{fp_text} object list or a
#' single \code{fp_text} objects. Use \code{fp_text(font.size = 0, ...)} to
#' inherit from default sizes of the presentation.
#' @export
#' @examples
#' library(magrittr)
#' pptx <- read_pptx()
#' pptx <- add_slide(x = pptx, layout = "Title and Content", master = "Office Theme")
#' pptx <- ph_with_text(x = pptx, type = "title", str = "Example title")
#' pptx <- ph_with_ul(
#'   x = pptx, type = "body", index = 1,
#'   str_list = c("Level1", "Level2", "Level2", "Level3", "Level3", "Level1"),
#'   level_list = c(1, 2, 2, 3, 3, 1),
#'   style = fp_text(color = "red", font.size = 0) )
#' print(pptx, target = tempfile(fileext = ".pptx"))
ph_with_ul <- function(x, type = "body", index = 1, str_list = character(0), level_list = integer(0),
                       style = NULL) {
  stopifnot(is.character(str_list))
  stopifnot(is.numeric(level_list))

  stopifnot(type %in% c("ctrTitle", "subTitle", "dt", "ftr",
                        "sldNum", "title", "body"))
  if (length(str_list) != length(level_list) & length(str_list) > 0) {
    stop("str_list and level_list have different lenghts.")
  }

  if( !is.null(style)){
    if( inherits(style, "fp_text") )
      style <- lapply(seq_len(length(str_list)), function(x) style )
    style_str <- sapply(style, format, type = "pml")
    style_str <- rep_len(style_str, length.out = length(str_list))
  } else style_str <- rep("<a:rPr/>", length(str_list))

  slide <- x$slide$get_slide(x$cursor)
  sh_pr_df <- slide$get_xfrm(type = type, index = index)

  tmpl <- "<a:p><a:pPr%s/><a:r>%s<a:t>%s</a:t></a:r></a:p>"
  lvl <- sprintf(" lvl=\"%.0f\"", level_list - 1)
  lvl <- ifelse(level_list > 1, lvl, "")
  p <- sprintf(tmpl, lvl, style_str, htmlEscape(str_list) )
  p <- paste(p, collapse = "")

  sh_pr_df$str <- p
  xml_elt <- do.call(pml_shape_par, sh_pr_df)
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
}

fortify_pml_images <- function(x, str){

  slide <- x$slide$get_slide(x$cursor)
  ref <- slide$rel_df()

  ref <- ref[ref$ext_src != "",]
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
  node <- xml_find_first( doc, as_xpath_content_sel("//") )
  node <- set_xfrm_attr(node, offx = xfrm$offx, offy = xfrm$offy,
                        cx = xfrm$cx, cy = xfrm$cy)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), doc)

  slide$fortify_id()
  x
}


#' @export
#' @rdname ph_from_xml
#' @param left,top location of the new shape on the slide
#' @param width,height shape size in inches
ph_from_xml_at <- function( x, value, left, top, width, height ){

  slide <- x$slide$get_slide(x$cursor)

  doc <- as_xml_document(value)

  node <- xml_find_first( doc, as_xpath_content_sel("//") )
  node <- set_xfrm_attr(node,
                        offx = left*914400,
                        offy = top*914400,
                        cx = width*914400,
                        cy = height*914400)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), doc)

  slide$fortify_id()
  x
}



