#' @export
#' @title add object into a new shape
#' @description add object into a new shape in a slide.
#' @param x a pptx device
#' @param value object to add as a new shape.
#' @param location a placeholder location object. This is a convenient
#' argument that can replace usage of arguments \code{type} and \code{index}.
#' See section \code{see also}.
#' @param ... Arguments to be passed to methods
#' @examples
#' library(magrittr)
#'
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Two Content",
#'   master = "Office Theme")
#' doc <- ph_with(x = doc, value = c("Un titre", "Deux titre"),
#'                location = ph_location_left() )
#' doc <- ph_with(x = doc, value = iris[1:4, 3:5],
#'                location = ph_location_right() )
#'
#' if( require("ggplot2") ){
#'   doc <- add_slide(doc, layout = "Title and Content",
#'     master = "Office Theme")
#'   gg_plot <- ggplot(data = iris ) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length),
#'       size = 3) +
#'     theme_minimal()
#'   doc <- ph_with(x = doc, value = gg_plot,
#'                  location = ph_location_fullsize() )
#' }
#'
#' doc <- add_slide(doc, layout = "Title and Content",
#'   master = "Office Theme")
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' doc <- ph_with(x = doc, external_img(img.file, 100/72, 76/72),
#'                location = ph_location_right(), use_loc_size = FALSE )
#'
#' svg_file <- file.path(R.home(component = "doc"), "html/Rlogo.svg")
#' if( require("rsvg") ){
#'   doc <- ph_with(x = doc, external_img(svg_file),
#'     location = ph_location_left(),
#'     use_loc_size = TRUE )
#' }
#'
#'
#'
#' # unordered_list ----
#' ul <- unordered_list(
#'   level_list = c(1, 2, 2, 3, 3, 1),
#'   str_list = c("Level1", "Level2", "Level2", "Level3", "Level3", "Level1"),
#'   style = fp_text(color = "red", font.size = 0) )
#' doc <- add_slide(doc, layout = "Title and Content",
#'                  master = "Office Theme")
#' doc <- ph_with(x = doc, value = ul,
#'                location = ph_location_fullsize() )
#'
#' print(doc, target = fileout )
#' @seealso [ph_location_type], [ph_location], [ph_location_label],
#' [ph_location_left], [ph_location_right], [ph_location_fullsize]
#' @importFrom rlang eval_tidy enquo call_modify
ph_with <- function(x, value, ...){
  UseMethod("ph_with", value)
}

gen_ph_str <- function( left = 0, top = 0, width = 3, height = 3,
                bg = "transparent", rot = 0, label = "", ph = "<p:ph/>"){

  if( is.null(rot)) rot <- 0

  if( !is.null(bg) && !is.color( bg ) )
    stop("bg must be a valid color.", call. = FALSE )

  bg_str <- ""
  if( !is.null(bg)){
    bg_str <- sprintf("<a:solidFill><a:srgbClr val=\"%s\"><a:alpha val=\"%.0f\"/></a:srgbClr></a:solidFill>",
            colcode0(bg), colalpha(bg) )
  }

  if( is.na(ph)){
    xfrm_str <- "<a:xfrm rot=\"%.0f\"><a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm>"
    xfrm_str <- sprintf(xfrm_str, -rot * 60000,
                        left * 914400, top * 914400,
                        width * 914400, height * 914400)
    ph = "<p:ph/>"
  } else {
    xfrm_str <- ""
  }


  str <- "<p:nvSpPr><p:cNvPr id=\"0\" name=\"%s\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr>%s%s</p:spPr>"
  sprintf(str, label, ph, xfrm_str, bg_str )

}
gen_ph_str_old <- function( left = 0, top = 0, width = 3, height = 3,
                bg = "transparent", rot = 0, label = "", ph = "<p:ph/>"){

  if( is.null(rot)) rot <- 0

  if( !is.null(bg) && !is.color( bg ) )
    stop("bg must be a valid color.", call. = FALSE )

  bg_str <- ""
  if( !is.null(bg)){
    bg_str <- sprintf("<a:solidFill><a:srgbClr val=\"%s\"><a:alpha val=\"%.0f\"/></a:srgbClr></a:solidFill>",
            colcode0(bg), colalpha(bg) )
  }

  xfrm_str <- "<a:xfrm rot=\"%.0f\"><a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm>"
  xfrm_str <- sprintf(xfrm_str, -rot * 60000,
                      left * 914400, top * 914400,
                      width * 914400, height * 914400)


  str <- "<p:nvSpPr><p:cNvPr id=\"0\" name=\"%s\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr>%s%s</p:spPr>"
  sprintf(str, label, ph, xfrm_str, bg_str )

}

#' @export
#' @title eval a location on the current slide
#' @description Eval a shape location against the current slide.
#' This function is to be used to add custom openxml code.
#' @param location a location for a placeholder. It has to be a
#' quosure.
#' @param x an rpptx object
#' @examples
#' library(rlang)
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content",
#'   master = "Office Theme")
#' location_eval(quo(ph_location_fullsize()), doc)
#' @seealso \code{\link{ph_location}}, \code{\link{ph_with}}
location_eval <- function(location, x){
  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm()
  location <- rlang::call_modify(location, x = x,
                             layout = unique( xfrm$name ),
                             master = unique(xfrm$master_name)
  )
  eval_tidy(location, data = x)
}

#' @export
#' @section with character:
#' When value is a character vector, each value will be
#' added as a paragraph.
#' @rdname ph_with
ph_with.character <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)

  location <- location_eval(rlang::enquo(location), x)

  new_ph <- gen_ph_str(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)
  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscape(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph,
                     "<p:txBody><a:bodyPr/><a:lstStyle/>",
                     pars, "</p:txBody></p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
}

#' @export
#' @param format_fun format function for non character vectors
#' @section with vector:
#' When value is a vector, the values will be first formatted and
#' then add as text, each value will be added as a paragraph.
#' @rdname ph_with
ph_with.numeric <- function(x, value, location, format_fun = format, ...){
  slide <- x$slide$get_slide(x$cursor)
  value <- format_fun(value, ...)
  location <- location_eval(rlang::enquo(location), x)

  new_ph <- gen_ph_str(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscape(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph,
                     "<p:txBody><a:bodyPr/><a:lstStyle/>",
                     pars, "</p:txBody></p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
}

#' @export
#' @rdname ph_with
ph_with.factor <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)
  value <- as.character(value)
  location <- location_eval(rlang::enquo(location), x)

  new_ph <- gen_ph_str(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscape(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph,
                     "<p:txBody><a:bodyPr/><a:lstStyle/>",
                     pars, "</p:txBody></p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
}
#' @export
#' @rdname ph_with
ph_with.logical <- ph_with.numeric

#' @export
#' @section with block_list:
#' When value is a block_list object, each value will be
#' added as a paragraph.
#' @rdname ph_with
ph_with.block_list <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)

  location <- location_eval(rlang::enquo(location), x)

  pars <- sapply(value, format, type = "pml")
  pars <- paste0(pars, collapse = "")

  new_ph <- gen_ph_str(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph,
                     "<p:txBody><a:bodyPr/><a:lstStyle/>",
                     pars, "</p:txBody></p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
}


#' @export
#' @section with unordered_list:
#' When value is a unordered_list object, each value will be
#' added as a paragraph.
#' @rdname ph_with
ph_with.unordered_list <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)
  location <- location_eval(rlang::enquo(location), x)

  if( !is.null(value$style)){
    style_str <- sapply(value$style, format, type = "pml")
    style_str <- rep_len(style_str, length.out = length(value$str))
  } else style_str <- rep("<a:rPr/>", length(value$str))
  tmpl <- "<a:p><a:pPr%s/><a:r>%s<a:t>%s</a:t></a:r></a:p>"
  lvl <- sprintf(" lvl=\"%.0f\"", value$lvl - 1)
  lvl <- ifelse(value$lvl > 1, lvl, "")
  p <- sprintf(tmpl, lvl, style_str, htmlEscape(value$str) )
  p <- paste(p, collapse = "")

  new_ph <- gen_ph_str(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph,
          "<p:txBody><a:bodyPr/><a:lstStyle/>", p, "</p:txBody></p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
}


#' @export
#' @section with data.frame:
#' When value is a data.frame, a simple table
#' is added, use package \code{flextable} instead
#' for more advanced formattings.
#' @param header display header if TRUE
#' @param first_row,last_row,first_column,last_column logical for PowerPoint table options
#' @rdname ph_with
ph_with.data.frame <- function(x, value, location, header = TRUE,
                               first_row = TRUE, first_column = FALSE,
                               last_row = FALSE, last_column = FALSE, ...){
  location <- location_eval(rlang::enquo(location), x)

  slide <- x$slide$get_slide(x$cursor)
  xml_elt <- table_shape(x = x, value = value, left = location$left*914400, top = location$top*914400,
                         width = location$width*914400, height = location$height*914400,
                         first_row = first_row, first_column = first_column,
                         last_row = last_row, last_column = last_column,
                         header = header )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))
  slide$fortify_id()
  x
}


#' @export
#' @section with a ggplot object:
#' When value is a ggplot object, a raster plot
#' is produced and added, use package \code{rvg}
#' instead for more advanced graphical features.
#' @rdname ph_with
ph_with.gg <- function(x, value, location, ...){
  location <- location_eval(rlang::enquo(location), x)

  slide <- x$slide$get_slide(x$cursor)
  if( !requireNamespace("ggplot2") )
    stop("package ggplot2 is required to use this function")

  slide <- x$slide$get_slide(x$cursor)
  width <- location$width
  height <- location$height

  stopifnot(inherits(value, "gg") )
  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')
  png(filename = file, width = width, height = height, units = "in", res = 300, ...)
  print(value)
  dev.off()
  on.exit(unlink(file))

  ext_img <- external_img(file, width = width, height = height)
  xml_elt <- format(ext_img, type = "pml")
  slide$reference_img(src = file, dir_name = file.path(x$package_dir, "ppt/media"))
  xml_elt <- fortify_pml_images(x, xml_elt)

  doc <- as_xml_document(xml_elt)

  node <- xml_find_first( doc, "p:spPr")
  off <- xml_child(node, "a:xfrm/a:off")
  xml_attr( off, "x") <- sprintf( "%.0f", location$left * 914400 )
  xml_attr( off, "y") <- sprintf( "%.0f", location$top * 914400 )

  xmlslide <- slide$get()

  xml_add_child(xml_find_first(xmlslide, "p:cSld/p:spTree"), doc)
  slide$fortify_id()
  x

}


#' @export
#' @section with a external_img object:
#' When value is a external_img object, image will be copied
#' into the PowerPoint presentation. The width and height
#' specified in call to \code{\link{external_img}} will be
#' ignored, their values will be those of the location,
#' unless use_loc_size is set to FALSE.
#' @param use_loc_size if set to FALSE, external_img width and height will
#' be used.
#' @rdname ph_with
ph_with.external_img <- function(x, value, location, use_loc_size = TRUE, ...){
  location <- location_eval(rlang::enquo(location), x)

  slide <- x$slide$get_slide(x$cursor)

  if( use_loc_size ){
    width <- location$width
    height <- location$height
  } else {
    width <- attr(value, "dims")$width
    height <- attr(value, "dims")$height
  }

  file_type <- gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", value)

  if( file_type %in% ".svg" ){
    if (!requireNamespace("rsvg")){
      stop("package 'rsvg' is required to convert svg file to rasters")
    }

    file <- tempfile(fileext = ".png")
    rsvg::rsvg_png(as.character(value), file = file)
    value[1] <- file
    file_type <- ".png"
  }

  new_src <- tempfile( fileext = gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", as.character(value)) )
  file.copy( as.character(value), to = new_src )
  ext_img <- external_img(new_src, width = width, height = height)
  xml_elt <- format(ext_img, type = "pml")
  slide$reference_img(src = new_src, dir_name = file.path(x$package_dir, "ppt/media"))
  xml_elt <- fortify_pml_images(x, xml_elt)

  doc <- as_xml_document(xml_elt)

  node <- xml_find_first( doc, "p:spPr")
  off <- xml_child(node, "a:xfrm/a:off")
  xml_attr( off, "x") <- sprintf( "%.0f", location$left * 914400 )
  xml_attr( off, "y") <- sprintf( "%.0f", location$top * 914400 )

  xmlslide <- slide$get()

  xml_add_child(xml_find_first(xmlslide, "p:cSld/p:spTree"), doc)
  slide$fortify_id()
  x

}




#' @export
ph_with.fpar <- function( x, value, location,
                          template_type = NULL, template_index = 1, ... ){

  location <- location_eval(rlang::enquo(location), x)

  slide <- x$slide$get_slide(x$cursor)

  new_ph <- gen_ph_str(left = location$left, top = location$top,
               width = location$width, height = location$height, label = location$label,
               rot = location$rotation, bg = location$bg, ph = location$ph)
  new_ph <- paste0( pml_with_ns("p:sp"), new_ph,"</p:sp>")
  if( !is.null( template_type ) ){
    stopifnot( template_type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )
    xfrm_df <- slide$get_xfrm(type = template_type, index = template_index)
    new_ph <- gsub("<p:ph/>", xfrm_df$ph, new_ph)
  }

  new_node <- as_xml_document(new_ph)

  p_ <- format( value, type = "pml")

  simple_shape <- paste0( pml_with_ns("p:txBody"),
                          "<a:bodyPr/><a:lstStyle/>",
                          p_, "</p:txBody>")
  xml_add_child(new_node, as_xml_document(simple_shape) )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), new_node)
  slide$fortify_id()$save()
  x
}




#' @export
#' @title add a new empty shape
#' @description add a new empty shape in the current slide.
#' @param x a pptx device
#' @param location a placeholder location object. This is a convenient
#' argument that can replace usage of arguments \code{type} and \code{index}.
#' See [ph_location_type], [ph_location], [ph_location_label],
#' [ph_location_left], [ph_location_right], [ph_location_fullsize].
#' @param type placeholder type (i.e. 'body', 'title')
#' @param index placeholder index (integer). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body', the first
#' one will be added with index 1 and the second one with index 2.
#' It is recommanded to use argument \code{location} instead of \code{type} and
#' \code{index}.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_empty(x = doc, type = "body", index = 1)
#' doc <- ph_empty(x = doc, location = ph_location_right())
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
ph_empty <- function( x, type = "body", index = 1, location = NULL ){

  stopifnot( type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body") )
  slide <- x$slide$get_slide(x$cursor)

  if( !is.null( rlang::get_expr(rlang::enquo(location)) ) ){
    location <- location_eval(rlang::enquo(location), x)
  } else {
    slide <- x$slide$get_slide(x$cursor)
    sh_pr_df <- slide$get_location(type = type, index = index)
    location <- ph_location(ph = sh_pr_df$ph, label = sh_pr_df$ph_label,
                            left = sh_pr_df$left, top = sh_pr_df$top,
                            width = sh_pr_df$width, height = sh_pr_df$height)
  }

  new_ph <- gen_ph_str(left = location$left, top = location$top,
                       width = location$width, height = location$height,
                       label = location$ph_label, ph = location$ph,
                       rot = location$rotation, bg = location$bg)
  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph, "</p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  slide$fortify_id()
  x
}


#' @export
#' @section with xml_document:
#' When value is an xml_document object, its content will be
#' added as a new shape in the current slide. This function
#' is to be used to add custom openxml code.
#' @rdname ph_with
ph_with.xml_document <- function( x, value, location, ... ){
  slide <- x$slide$get_slide(x$cursor)

  location <- location_eval(rlang::enquo(location), x)
  node <- xml_find_first( value, as_xpath_content_sel("//") )
  node <- set_xfrm_attr(node, offx = location$left*914400, offy = location$top*914400,
                        cx = location$width*914400, cy = location$height*914400)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)

  slide$fortify_id()
  x
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


# old functions -----------

#' @export
#' @title add text into a new shape
#' @description add text into a new shape in a slide.
#' This function will be deprecated in favor of \code{\link{ph_with}}
#' in the next release.
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
#' @inherit ph_empty seealso
ph_with_text <- function( x, str, type = "title", index = 1, location = NULL ){

  if(!is.character(str))
    str <- format(str)

  if( !is.null( rlang::get_expr(rlang::enquo(location)) ) ){
    ph_with(x, external_img(src=src, width = width, height = height), location = !!rlang::enexpr( location ))
  } else {
    slide <- x$slide$get_slide(x$cursor)
    sh_pr_df <- slide$get_location(type = type, index = index)
    ph_with(x, str, location = ph_location(ph = sh_pr_df$ph, label = sh_pr_df$ph_label, left = sh_pr_df$left, top = sh_pr_df$top,
                                           width = sh_pr_df$width, height = sh_pr_df$height))
  }
}


#' @export
#' @title add table
#' @description add a table as a new shape in the current slide.
#' This function will be deprecated in favor of \code{\link{ph_with}}
#' in the next release.
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
                           last_row = FALSE, last_column = FALSE,
                           location = NULL ){
  stopifnot(is.data.frame(value))

  if( !is.null( rlang::get_expr(rlang::enquo(location)) ) ){
    ph_with(x, value, location = !!rlang::enexpr( location ), header = header,
            first_row = first_row, first_column = first_column,
            last_row = last_row, last_column = last_column)
  } else {
    slide <- x$slide$get_slide(x$cursor)
    sh_pr_df <- slide$get_location(type = type, index = index)
    ph_with(x, value, location = ph_location(ph = sh_pr_df$ph, label = sh_pr_df$ph_label, left = sh_pr_df$left, top = sh_pr_df$top,
                                             width = sh_pr_df$width, height = sh_pr_df$height),
            header = header,
            first_row = first_row, first_column = first_column,
            last_row = last_row, last_column = last_column)
  }
}



#' @export
#' @title add image
#' @description add an image as a new shape in the current slide.
#' This function will be deprecated in favor of \code{\link{ph_with}}
#' in the next release.
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
ph_with_img <- function( x, src, type = "body", index = 1,
                         width = NULL, height = NULL,
                         location = NULL ){

  if( !is.null( rlang::get_expr(rlang::enquo(location)) ) )
    ph_with(x, external_img(src=src, width = width, height = height), location = !!rlang::enexpr( location ))
  else {
    slide <- x$slide$get_slide(x$cursor)
    sh_pr_df <- slide$get_location(type = type, index = index)
    if( is.null( width ) ) width <- sh_pr_df$width
    if( is.null( height ) ) height <- sh_pr_df$height
    ph_with(x, external_img(src=src, width = width, height = height),
            location = ph_location(ph = sh_pr_df$ph, label = sh_pr_df$ph_label, left = sh_pr_df$left, top = sh_pr_df$top,
                                   width = width, height = height))
  }
}

#' @export
#' @title add ggplot to a pptx presentation
#' @description add a ggplot as a png image into an rpptx object
#' This function will be deprecated in favor of \code{\link{ph_with}}
#' in the next release.
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
ph_with_gg <- function( x, value, type = "body", index = 1,
                        width = NULL, height = NULL, location = NULL, ... ){


  if( !is.null( rlang::get_expr(rlang::enquo(location)) ) )
    ph_with(x, value, location = !!rlang::enexpr( location ))
  else {
    slide <- x$slide$get_slide(x$cursor)
    sh_pr_df <- slide$get_location(type = type, index = index)
    if( is.null( width ) ) width <- sh_pr_df$width
    if( is.null( height ) ) height <- sh_pr_df$height
    ph_with(x, value, location = ph_location(ph = sh_pr_df$ph, label = sh_pr_df$ph_label, left = sh_pr_df$left, top = sh_pr_df$top,
                                             width = width, height = height))
  }

}

#' @title add unordered list to a pptx presentation
#' @description add an unordered list of text
#' into an rpptx object. Each text is associated with
#' a hierarchy level.
#' This function will be deprecated in favor of \code{\link{ph_with}}
#' in the next release.
#' @inheritParams ph_empty
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
#'   x = pptx, location = ph_location_type(type="body"),
#'   str_list = c("Level1", "Level2", "Level2", "Level3", "Level3", "Level1"),
#'   level_list = c(1, 2, 2, 3, 3, 1),
#'   style = fp_text(color = "red", font.size = 0) )
#' print(pptx, target = tempfile(fileext = ".pptx"))
ph_with_ul <- function(x, type = "body", index = 1,
                       str_list = character(0), level_list = integer(0),
                       style = NULL,
                       location = NULL) {
  value <- unordered_list(
    level_list = level_list,
    str_list = str_list,
    style = style )

  if( !is.null( rlang::get_expr(rlang::enquo(location)) ) ){
    ph_with(x = x, value = value, location = !!rlang::enexpr( location ) )
  } else {
    slide <- x$slide$get_slide(x$cursor)
    sh_pr_df <- slide$get_location(type = type, index = index)
    ph_with(x, value, location = ph_location(ph = sh_pr_df$ph, label = sh_pr_df$ph_label, left = sh_pr_df$left, top = sh_pr_df$top,
                                             width = sh_pr_df$width, height = sh_pr_df$height))
  }
}



# old functions_at -----------

#' @rdname ph_empty
#' @export
#' @param left,top location of the new shape on the slide
#' @param width,height shape size in inches
#' @param bg background color
#' @param rot rotation angle
#' @param template_type placeholder template type. If used, the new shape will
#' inherit the style from the placeholder template. If not used, no text
#' property is defined and for example text lists will not be indented.
#' @param template_index placeholder template index (integer). To be used when a placeholder
#' template type is not unique in the current slide, e.g. two placeholders with
#' type 'body'.
#' @examples
#'
#' # demo ph_empty_at ------
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_empty_at(x = doc, left = 1, top = 2, width = 5, height = 4)
#'
#' print(doc, target = fileout )
ph_empty_at <- function( x, left, top, width, height, bg = "transparent", rot = 0,
                         template_type = NULL, template_index = 1 ){

  ph <- NA_character_
  label <- ""
  type <- "body"
  slide <- x$slide$get_slide(x$cursor)

  if( !is.null( template_type ) ){
    xfrm_df <- slide$get_xfrm(type = template_type, index = template_index)
    ph <- xfrm_df$ph
  } else {
    ph <- sprintf('<p:ph type="%s"/>', type)
  }
  new_ph <- gen_ph_str_old(left = left, top = top,
                           width = width, height = height,
                           label = "", ph = ph,
                           rot = rot, bg = bg)


  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph, "</p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  slide$fortify_id()
  x

}


#' @export
#' @rdname ph_with_img
#' @param left,top location of the new shape on the slide
#' @param rot rotation angle
#' @examples
#'
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- ph_with_img_at(x = doc, src = img.file, height = 1.06, width = 1.39,
#'     left = 4, top = 4, rot = 45 )
#' }
#'
#' print(doc, target = fileout )
ph_with_img_at <- function( x, src, left, top, width, height, rot = 0 ){

  ph_with(x, external_img(src=src, width = width, height = height),
          location = ph_location(ph = "", label = "", left = left, top = top, width = width, height = height))

}

#' @export
#' @rdname ph_with_table
#' @param left,top location of the new shape on the slide
#' @param width,height shape size in inches
#' @examples
#'
#' library(magrittr)
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with_table_at(value = mtcars[1:6,],
#'     height = 4, width = 8, left = 4, top = 4,
#'     last_row = FALSE, last_column = FALSE, first_row = TRUE)
#'
#' print(doc, target = tempfile(fileext = ".pptx"))
ph_with_table_at <- function( x, value, left, top, width, height,
                              header = TRUE,
                              first_row = TRUE, first_column = FALSE,
                              last_row = FALSE, last_column = FALSE ){

  stopifnot(is.data.frame(value))
  ph_with(x,
          value,
          location = ph_location(
            ph = "", label = "", left = left, top = top,
            width = width, height = height),
          header = header,
          first_row = first_row, first_column = first_column,
          last_row = last_row, last_column = last_column)
}

#' @export
#' @param left,top location of the new shape on the slide
#' @importFrom grDevices png dev.off
#' @rdname ph_with_gg
ph_with_gg_at <- function( x, value, width, height, left, top, ... ){

  if( !requireNamespace("ggplot2") )
    stop("package ggplot2 is required to use this function")

  stopifnot(inherits(value, "gg"))

  ph_with(x,
          value,
          location = ph_location(
            ph = "", label = "", left = left, top = top,
            width = width, height = height), ...)
}






#' @export
#' @title add multiple formated paragraphs
#' @description add several formated paragraphs in a new shape in the current slide.
#' @param x rpptx object
#' @param fpars list of \code{\link{fpar}} objects
#' @param fp_pars list of \code{\link{fp_par}} objects. The list can contain
#' NULL to keep defaults.
#' @param left,top location of the new shape on the slide
#' @param width,height shape size in inches
#' @param bg background color
#' @param rot rotation angle
#' @param template_type placeholder template type. If used, the new shape will
#' inherit the style from the placeholder template. If not used, no text
#' property is defined and for example text lists will not be indented.
#' @param template_index placeholder template index (integer). To be used when a placeholder
#' template type is not unique in the current slide, e.g. two placeholders with
#' type 'body'.
#' @examples
#'
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content",
#'   master = "Office Theme")
#'
#' bold_face <- shortcuts$fp_bold(font.size = 0)
#' bold_redface <- update(bold_face, color = "red")
#'
#' fpar_1 <- fpar(
#'   ftext("Hello ", prop = bold_face), ftext("World", prop = bold_redface ),
#'   ftext(", \r\nhow are you?", prop = bold_face ) )
#'
#' fpar_2 <- fpar(
#'   ftext("Hello ", prop = bold_face), ftext("World", prop = bold_redface ),
#'   ftext(", \r\nhow are you again?", prop = bold_face ) )
#'
#' doc <- ph_with_fpars_at(x = doc, fpars = list(fpar_1, fpar_2),
#'   fp_pars = list(NULL, fp_par(text.align = "center")),
#'   left = 1, top = 2, width = 7, height = 4)
#' doc <- ph_with_fpars_at(x = doc, fpars = list(fpar_1, fpar_2),
#'   template_type = "body", template_index = 1,
#'   left = 4, top = 5, width = 4, height = 3)
#'
#' print(doc, target = fileout )
ph_with_fpars_at <- function( x, fpars = list(), fp_pars = list(),
                              left, top, width, height, bg = "transparent", rot = 0,
                              template_type = NULL, template_index = 1 ){

  if( length(fp_pars) < 1 )
    fp_pars <- lapply(fpars, function(x) NULL )
  if( length(fp_pars) != length(fpars) )
    stop("fp_pars and fpars should have the same length")

  p_ <- mapply(
    function(fpar, fp_par) {
      if( !is.null(fp_par) ) {
        fpar <- update(fpar, fp_p = fp_par)
      }
      fpar
    },
    fpar = fpars, fp_par = fp_pars, SIMPLIFY = FALSE )
  p_ <- do.call(block_list, p_)

  ph <- ""
  label <- ""
  if( !is.null( template_type ) ){
    slide <- x$slide$get_slide(x$cursor)
    xfrm_df <- slide$get_xfrm(type = template_type, index = template_index)
    ph <- xfrm_df$ph
    label <- xfrm_df$ph_label
  }

  ph_with(x, p_,
          location = ph_location(
            ph = ph, label = label, left = left, top = top,
            width = width, height = height))

}


