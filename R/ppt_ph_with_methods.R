#' @export
#' @title add objects into a new shape on the current slide
#' @description add object into a new shape in the current slide. This
#' function is able to add all supported outputs to a presentation
#' and should replace calls to older functions starting with
#' \code{ph_with_*}.
#' @param x an rpptx object
#' @param value object to add as a new shape. Supported objects
#' are vectors, data.frame, graphics, block of formatted paragraphs,
#' unordered list of formatted paragraphs,
#' pretty tables with package flextable, editable graphics with
#' package rvg, 'Microsoft' charts with package mschart.
#' @param location a placeholder location object.
#' It will be used to specify the location of the new shape. This location
#' can be defined with a call to one of the ph_location functions. See
#' section \code{see also}.
#' @param ... Arguments to be passed to methods
#' @examples
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
#'   doc <- add_slide(doc)
#'   gg_plot <- ggplot(data = iris ) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length),
#'       size = 3) +
#'     theme_minimal()
#'   doc <- ph_with(x = doc, value = gg_plot,
#'                  location = ph_location_fullsize() )
#'   doc <- ph_with(x = doc, value = "graphic title",
#'                location = ph_location_type(type="title") )
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
#' # block list ------
#' bl <- block_list(
#'   fpar(ftext("hello world", shortcuts$fp_bold(color = "pink"))),
#'   fpar(
#'     ftext("hello", shortcuts$fp_bold()),
#'     ftext("hello", shortcuts$fp_italic(color="red"))
#'   ))
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, value = bl,
#'                location = ph_location_type(type="body") )
#'
#' # fpar ------
#' hw <- fpar(ftext("hello world", shortcuts$fp_bold(color = "pink")))
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, value = hw,
#'                location = ph_location_type(type="body") )


#' # unordered_list ----
#' ul <- unordered_list(
#'   level_list = c(1, 2, 2, 3, 3, 1),
#'   str_list = c("Level1", "Level2", "Level2", "Level3", "Level3", "Level1"),
#'   style = fp_text(color = "red", font.size = 0) )
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, value = ul,
#'                location = ph_location_fullsize() )
#'
#' print(doc, target = fileout )
#' @seealso [ph_location_type], [ph_location], [ph_location_label],
#' [ph_location_left], [ph_location_right], [ph_location_fullsize],
#' [ph_location_template]
#' @importFrom rlang eval_tidy
ph_with <- function(x, value, ...){
  UseMethod("ph_with", value)
}

gen_bg_str <- function(bg){
  bg_str <- ""
  if( !is.null(bg)){
    bg_str <- sprintf("%s<a:srgbClr val=\"%s\"><a:alpha val=\"%.0f\"/></a:srgbClr></a:solidFill>",
                      pml_with_ns("a:solidFill") , colcode0(bg), colalpha(bg) )
  }
  bg_str
}


gen_ph_str <- function( left = 0, top = 0, width = 3, height = 3,
                bg = "transparent", rot = 0, label = "", ph = "<p:ph/>"){

  if( is.null(rot)) rot <- 0

  if( !is.null(bg) && !is.color( bg ) )
    stop("bg must be a valid color.", call. = FALSE )

  bg_str <- gen_bg_str(bg)

  xfrm_str <- "<a:xfrm rot=\"%.0f\"><a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm>"
  xfrm_str <- sprintf(xfrm_str, -rot * 60000,
                      left * 914400, top * 914400,
                      width * 914400, height * 914400)
  if( is.null(ph) || is.na(ph)){
    ph = "<p:ph/>"
  }


  str <- "<p:nvSpPr><p:cNvPr id=\"0\" name=\"%s\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr>%s%s</p:spPr>"
  sprintf(str, label, ph, xfrm_str, bg_str )

}
#' @export
#' @title Utility to eval a location
#' @description Eval a shape location with fortify_location.
#' This function will be removed in the next release; it was
#' required when location was a quosure but this is no more
#' necessary.
#' @param location a location for a placeholder.
#' @param x an rpptx object
#' @seealso \code{\link{ph_location}}, \code{\link{ph_with}}
location_eval <- function(location, x){
  if(inherits(location, "quosure")){
    fortify_location(eval_tidy(location), x)
  } else {
    fortify_location(location, x)
  }
}

#' @export
#' @describeIn ph_with add a character vector to a new shape on the
#' current slide, values will be added as paragraphs.
ph_with.character <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)
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
#' @describeIn ph_with add a numeric vector to a new shape on the
#' current slide, values will be be first formatted then
#' added as paragraphs.
ph_with.numeric <- function(x, value, location, format_fun = format, ...){
  slide <- x$slide$get_slide(x$cursor)
  value <- format_fun(value, ...)
  location <- fortify_location(location, doc = x)

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
#' @describeIn ph_with add a factor vector to a new shape on the
#' current slide, values will be be converted as character and then
#' added as paragraphs.
ph_with.factor <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)
  value <- as.character(value)
  location <- fortify_location(location, doc = x)

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
#' @describeIn ph_with add a \code{\link{block_list}} made
#' of \code{\link{fpar}} to a new shape on the current slide.
ph_with.block_list <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)

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
#' @describeIn ph_with add a \code{\link{unordered_list}} made
#' of \code{\link{fpar}} to a new shape on the current slide.
ph_with.unordered_list <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)
  location <- fortify_location(location, doc = x)

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
#' @param header display header if TRUE
#' @param first_row,last_row,first_column,last_column logical for PowerPoint table options
#' @describeIn ph_with add a data.frame to a new shape on the current slide. Use
#' package \code{flextable} instead for more advanced formattings.
ph_with.data.frame <- function(x, value, location, header = TRUE,
                               first_row = TRUE, first_column = FALSE,
                               last_row = FALSE, last_column = FALSE, ...){
  location <- fortify_location(location, doc = x)

  slide <- x$slide$get_slide(x$cursor)
  xml_elt <- table_shape(x = x, value = value, left = location$left*914400, top = location$top*914400,
                         width = location$width*914400, height = location$height*914400,
                         first_row = first_row, first_column = first_column,
                         last_row = last_row, last_column = last_column,
                         header = header )

  value <- as_xml_document(xml_elt)
  xml_to_slide(slide, location, value)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)
  slide$fortify_id()
  x
}


#' @export
#' @describeIn ph_with add a ggplot object to a new shape on the
#' current slide. Use package \code{rvg} for more advanced graphical features.
ph_with.gg <- function(x, value, location, ...){
  location <- fortify_location(location, doc = x)
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

  value <- as_xml_document(xml_elt)
  xml_to_slide(slide, location, value)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)
  slide$fortify_id()
  x

}


#' @export
#' @param use_loc_size if set to FALSE, external_img width and height will
#' be used.
#' @describeIn ph_with add a \code{\link{external_img}} to a new shape
#' on the current slide.
#'
#' When value is a external_img object, image will be copied
#' into the PowerPoint presentation. The width and height
#' specified in call to \code{external_img} will be
#' ignored, their values will be those of the location,
#' unless use_loc_size is set to FALSE.
ph_with.external_img <- function(x, value, location, use_loc_size = TRUE, ...){
  location <- fortify_location(location, doc = x)

  slide <- x$slide$get_slide(x$cursor)

  if( !use_loc_size ){
    location$width <- attr(value, "dims")$width
    location$height <- attr(value, "dims")$height
  }
  width <- location$width
  height <- location$height

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

  value <- as_xml_document(xml_elt)
  xml_to_slide(slide, location, value)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)
  slide$fortify_id()
  x

}




#' @export
#' @describeIn ph_with add an \code{\link{fpar}} to a new shape
#' on the current slide as a single paragraph in a \code{\link{block_list}}.
ph_with.fpar <- function( x, value, location, ... ){

  ph_with.block_list(x, value = block_list(value), location = location)

  x
}




#' @export
#' @title add a new empty shape
#' @description add a new empty shape in the current slide.
#' This function was implemented for development purpose and
#' should not be used.
#' @param x an pptx object
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
  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
  }
  new_ph <- gen_ph_str(left = 0, top = 0,
                       width = 3, height = 3)
  xml_elt <- paste0( pml_with_ns("p:sp"), new_ph,
                     "<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody></p:sp>" )
  node <- as_xml_document(xml_elt)

  ph_with(x, node, location = location)
}

xml_to_slide <- function(slide, location, value){
  node <- xml_find_first( value, as_xpath_content_sel("//") )
  if(xml_name(node) == "grpSp"){
    node_sppr <- xml_child(node, "p:grpSpPr")
    node_xfrm <- xml_child(node, "p:grpSpPr/a:xfrm")
    node_name <- xml_child(node, "p:nvGrpSpPr/p:cNvPr")
  } else if(xml_name(node) == "graphicFrame") {
    node_xfrm <- xml_child(node, "p:xfrm")
    node_name <- xml_child(node, "p:nvGraphicFramePr/p:cNvPr")
    node_sppr <- xml_missing()
  } else if(xml_name(node) == "pic"){
    node_sppr <- xml_child(node, "p:spPr")
    node_xfrm <- xml_child(node, "p:spPr/a:xfrm")
    node_name <- xml_child(node, "p:nvPicPr/p:cNvPr")
  } else if(xml_name(node) == "sp"){
    node_sppr <- xml_child(node, "p:spPr")
    node_xfrm <- xml_child(node, "p:spPr/a:xfrm")
    node_name <- xml_child(node, "p:nvSpPr/p:cNvPr")
  } else {
    node_xfrm <- xml_missing()
    node_name <- xml_missing()
    node_sppr <- xml_missing()
  }


  if( !inherits(node_name, "xml_missing") ){
    xml_attr( node_name, "name") <- location$ph_label
  }

  if( !inherits(node_xfrm, "xml_missing") ){
    off <- xml_child(node_xfrm, "a:off")
    ext <- xml_child(node_xfrm, "a:ext")
    chOff <- xml_child(node_xfrm, "a:chOff")
    chExt <- xml_child(node_xfrm, "a:chExt")
    xml_attr( off, "x") <- sprintf( "%.0f", location$left*914400 )
    xml_attr( off, "y") <- sprintf( "%.0f", location$top*914400 )
    xml_attr( ext, "cx") <- sprintf( "%.0f", location$width*914400 )
    xml_attr( ext, "cy") <- sprintf( "%.0f", location$height*914400 )
    xml_attr( chOff, "x") <- sprintf( "%.0f", location$left*914400 )
    xml_attr( chOff, "y") <- sprintf( "%.0f", location$top*914400 )
    xml_attr( chExt, "cx") <- sprintf( "%.0f", location$width*914400 )
    xml_attr( chExt, "cy") <- sprintf( "%.0f", location$height*914400 )

    # add location$rotation to cNvPr:name
    if( !is.null(location$rotation) && is.numeric(location$rotation) ) {
      xml_set_attr(node_xfrm, "rot", sprintf("%.0f", -location$rotation * 60000))
    }
  }

  if( !inherits(node_sppr, "xml_missing") ){
    # add location$bg to SpPr
    if( !is.null(location$bg) ) {
      bg_str <- gen_bg_str(location$bg)
      xml_add_child(node_sppr, as_xml_document(bg_str))
    }
  }
  value
}

#' @export
#' @importFrom xml2 xml_child xml_set_attr xml_missing
#' @describeIn ph_with add an xml_document object to a new shape on the
#' current slide. This function is to be used to add custom openxml code.
ph_with.xml_document <- function( x, value, location, ... ){
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)

  xml_to_slide(slide, location, value)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)

  slide$fortify_id()
  x
}

# old functions -----------

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

#' @export
#' @title add text into a new shape
#' @description add text into a new shape in a slide.
#' This function will be deprecated in favor of \code{\link{ph_with}}
#' in the next release.
#' @inheritParams ph_empty
#' @param str text to add
#' @examples
#' # define locations for placeholders ----
#' loc_title <- ph_location_type(type = "title")
#' loc_footer <- ph_location_type(type = "ftr")
#' loc_dt <- ph_location_type(type = "dt")
#' loc_slidenum <- ph_location_type(type = "sldNum")
#' loc_body <- ph_location_type(type = "body")
#'
#'
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, "Un titre", location = loc_title)
#' doc <- ph_with(x = doc, "pied de page", location = loc_footer)
#' doc <- ph_with(x = doc, format(Sys.Date()), location = loc_dt)
#' doc <- ph_with(x = doc, "slide 1", location = loc_slidenum)
#' doc <- ph_with(x = doc, letters[1:10], location = loc_body)
#'
#' loc_subtitle <- ph_location_type(type = "subTitle")
#' loc_ctrtitle <- ph_location_type(type = "ctrTitle")
#' doc <- add_slide(doc, layout = "Title Slide", master = "Office Theme")
#' doc <- ph_with(x = doc, "Un sous titre", location = loc_subtitle)
#' doc <- ph_with(x = doc, "Un titre", location = loc_ctrtitle)
#'
#' fileout <- tempfile(fileext = ".pptx")
#' print(doc, target = fileout )
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
#' @inherit ph_empty seealso
ph_with_text <- function( x, str, type = "title", index = 1, location = NULL ){

  if(!is.character(str))
    str <- format(str)
  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
  }
  ph_with(x, str, location = location)
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
ph_with_table <- function( x, value, type = "body", index = 1,
                           header = TRUE,
                           first_row = TRUE, first_column = FALSE,
                           last_row = FALSE, last_column = FALSE,
                           location = NULL ){
  .Deprecated(new = "ph_with")
  stopifnot(is.data.frame(value))

  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
  }
  ph_with(x, value, location = location, header = header,
          first_row = first_row, first_column = first_column,
          last_row = last_row, last_column = last_column)
}



#' @export
#' @title add image
#' @description add an image as a new shape in the current slide.
#' This function will be deprecated in favor of \code{\link{ph_with}}
#' in the next release.
#' @inheritParams ph_empty
#' @param src image filename, the basename of the file must not contain any blank.
#' @param width,height image size in inches
#' @importFrom xml2 xml_find_first as_xml_document xml_remove
ph_with_img <- function( x, src, type = "body", index = 1,
                         width = NULL, height = NULL,
                         location = NULL ){
  .Deprecated(new = "ph_with")
  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
  }
  ph_with(x, external_img(src=src, width = width, height = height), location = location)
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
ph_with_gg <- function( x, value, type = "body", index = 1,
                        width = NULL, height = NULL, location = NULL, ... ){
  .Deprecated(new = "ph_with")
  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
    if( !is.null( width ) ) location$width <- width
    if( !is.null( height ) ) location$height <- height
  }
  ph_with(x, value, location = location)

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
ph_with_ul <- function(x, type = "body", index = 1,
                       str_list = character(0), level_list = integer(0),
                       style = NULL,
                       location = NULL) {
  .Deprecated(new = "ph_with")
  value <- unordered_list(
    level_list = level_list,
    str_list = str_list,
    style = style )
  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
  }
  ph_with(x = x, value = value, location = location )
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
ph_empty_at <- function( x, left, top, width, height, bg = "transparent", rot = 0,
                         template_type = NULL, template_index = 1 ){
  .Deprecated(new = "ph_with")
  location <- ph_location_template(left = left, top = top, width = width, height = height,
              label = "", type = template_type, id = template_index)
  ph_with(x, "", location = location)
}


#' @export
#' @rdname ph_with_img
#' @param left,top location of the new shape on the slide
#' @param rot rotation angle
ph_with_img_at <- function( x, src, left, top, width, height, rot = 0 ){

  ph_with(x, external_img(src=src, width = width, height = height),
          location = ph_location(ph = "", label = "", left = left, top = top, width = width, height = height, rotation = rot))

}

#' @export
#' @rdname ph_with_table
#' @param left,top location of the new shape on the slide
#' @param width,height shape size in inches
ph_with_table_at <- function( x, value, left, top, width, height,
                              header = TRUE,
                              first_row = TRUE, first_column = FALSE,
                              last_row = FALSE, last_column = FALSE ){

  .Deprecated(new = "ph_with")
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
ph_with_fpars_at <- function( x, fpars = list(), fp_pars = list(),
                              left, top, width, height, bg = "transparent", rot = 0,
                              template_type = NULL, template_index = 1 ){

  .Deprecated(new = "ph_with")
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

  location <- ph_location_template(left = left, top = top, width = width, height = height,
                                   label = "", type = template_type, id = template_index)
  ph_with(x, p_, location = location)
}


