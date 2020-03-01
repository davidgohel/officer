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
#' section \code{"see also"}.
#' @param ... further arguments passed to or from other methods. When
#' adding a `ggplot` object or `plot_instr`, these arguments will be used
#' by png function.
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
#'
#' anyplot <- plot_instr(code = {
#'   barplot(1:5, col = 2:6)
#' })
#'
#' doc <- add_slide(doc)
#' doc <- ph_with(
#'   doc, anyplot,
#'   location = ph_location_fullsize(),
#'   bg = "#00000066", pointsize = 12)
#'
#' if( require("ggplot2") ){
#'   doc <- add_slide(doc)
#'   gg_plot <- ggplot(data = iris ) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length),
#'       size = 3) +
#'     theme_minimal()
#'   doc <- ph_with(x = doc, value = gg_plot,
#'                  location = ph_location_fullsize(),
#'                  bg = "transparent" )
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
ph_with <- function(x, value, location, ...){
  UseMethod("ph_with", value)
}


#' @export
#' @describeIn ph_with add a character vector to a new shape on the
#' current slide, values will be added as paragraphs.
ph_with.character <- function(x, value, location, ...){
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)
  new_ph <- sh_props_pml(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)
  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscapeCopy(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0( psp_ns_yes, new_ph,
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

  new_ph <- sh_props_pml(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscapeCopy(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0( psp_ns_yes, new_ph,
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

  new_ph <- sh_props_pml(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscapeCopy(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0( psp_ns_yes, new_ph,
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
#' @param is_list experimental paramater to make
#' block_list formated as an unordered list. This
#' should evolve in the next versions.
#' @describeIn ph_with add a \code{\link{block_list}} made
#' of \code{\link{fpar}} to a new shape on the current slide.
ph_with.block_list <- function(x, value, location, is_list = FALSE, ...){
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)

  pars <- sapply(value, to_pml)
  if( is_list ){
    pars <- gsub("<a:buNone/>", "", pars, fixed = TRUE)
  }

  pars <- paste0(pars, collapse = "")

  new_ph <- sh_props_pml(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  xml_elt <- paste0( psp_ns_yes, new_ph,
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

  p <- to_pml(value)

  new_ph <- sh_props_pml(left = location$left, top = location$top,
               width = location$width, height = location$height,
               label = location$ph_label, ph = location$ph,
               rot = location$rotation, bg = location$bg)

  xml_elt <- paste0( psp_ns_yes, new_ph,
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
  style_id <- x$table_styles$def[1]

  bt <- block_table(x = value, style = style_id, header = header, first_row = first_row,
              first_column = first_column, last_row = last_row,
              last_column = last_column)

  xml_elt <- to_pml(
    bt, left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg)

  value <- as_xml_document(xml_elt)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)
  slide$fortify_id()
  x
}


#' @export
#' @describeIn ph_with add a ggplot object to a new shape on the
#' current slide. Use package \code{rvg} for more advanced graphical features.
#' @param res resolution of the png image in ppi
ph_with.gg <- function(x, value, location, res = 300, ...){
  location_ <- fortify_location(location, doc = x)
  slide <- x$slide$get_slide(x$cursor)
  if( !requireNamespace("ggplot2") )
    stop("package ggplot2 is required to use this function")

  slide <- x$slide$get_slide(x$cursor)
  width <- location_$width
  height <- location_$height

  stopifnot(inherits(value, "gg") )
  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')
  png(filename = file, width = width, height = height, units = "in", res = res, ...)
  print(value)
  dev.off()
  on.exit(unlink(file))

  ext_img <- external_img(file, width = width, height = height)
  ph_with(x, ext_img, location = location )
}

#' @export
#' @describeIn ph_with add an R plot to a new shape on the
#' current slide. Use package \code{rvg} for more advanced graphical features.
ph_with.plot_instr <- function(x, value, location, res = 300, ...){
  location_ <- fortify_location(location, doc = x)
  slide <- x$slide$get_slide(x$cursor)
  slide <- x$slide$get_slide(x$cursor)
  width <- location_$width
  height <- location_$height

  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')

  dirname <- tempfile( )
  dir.create( dirname )
  filename <- paste( dirname, "/plot%03d.png" ,sep = "" )
  png(filename = filename, width = width, height = height, units = "in", res = res, ...)

  tryCatch({
    eval(value$code)
  },
  finally = {
    dev.off()
  } )
  file = list.files( dirname , full.names = TRUE )
  on.exit(unlink(dirname, recursive = TRUE, force = TRUE))

  if( length( file ) > 1 ){
    stop( length( file )," files have been produced. Multiple plot are not supported")
  }

  ext_img <- external_img(file, width = width, height = height)
  ph_with(x, ext_img, location = location)
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

  xml_str <- pic_pml(left = location$left, top = location$top,
                       width = width, height = height,
                       label = location$ph_label, ph = location$ph,
                       rot = location$rotation, bg = location$bg, src = new_src)

  slide$reference_img(src = new_src, dir_name = file.path(x$package_dir, "ppt/media"))
  xml_elt <- fortify_pml_images(x, xml_str)

  value <- as_xml_document(xml_elt)
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
#' @describeIn ph_with add an \code{\link{empty_content}} to a new shape
#' on the current slide.
ph_with.empty_content <- function( x, value, location, ... ){

  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)
  new_ph <- sh_props_pml(left = location$left, top = location$top,
                       width = location$width, height = location$height,
                       label = location$ph_label, ph = location$ph,
                       rot = location$rotation, bg = location$bg)
  xml_elt <- paste0( psp_ns_yes, new_ph, "</p:sp>" )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)

  slide$fortify_id()
  x
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
      bg_str <- solid_fill_pml(location$bg)
      xml_add_child(node_sppr, as_xml_document(bg_str))
    }
  }
  value
}

#' @export
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
