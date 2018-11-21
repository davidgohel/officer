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
#' doc <- add_slide(doc, layout = "Title and Content",
#'   master = "Office Theme")
#'
#' if( require("ggplot2") ){
#'   gg_plot <- ggplot(data = iris ) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length),
#'       size = 3) +
#'     theme_minimal()
#'   doc <- ph_with(x = doc, value = gg_plot,
#'                  location = ph_location_fullsize(doc) )
#' }
#'
#' doc <- ph_with(x = doc, value = c("Un titre", "Deux titre"),
#'                location = ph_location_left(doc) )
#' doc <- ph_with(x = doc, value = iris[1:4, 3:5],
#'                location = ph_location_right(doc) )
#'
#' print(doc, target = fileout )
#' @seealso [ph_location_type], [ph_location], [ph_location_label],
#' [ph_location_left], [ph_location_right], [ph_location_fullsize]
ph_with <- function(x, value, location, ...){
  UseMethod("ph_with", value)
}


#' @export
#' @section with character:
#' When value is a character vector, each value will be
#' added as a paragraph.
#' @rdname ph_with
ph_with.character <- function(x, value, location, ...){

  slide <- x$slide$get_slide(x$cursor)
  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscape(value), "</a:t></a:r></a:p>", collapse = "")
  sp_pr <- sprintf("<p:spPr><a:xfrm rot=\"%.0f\"><a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm></p:spPr>",
                   ifelse(is.null(location$rotation), 0, -location$rotation) * 60000,
                   location$left * 914400, location$top * 914400,
                   location$width * 914400, location$height * 914400)

  nv_sp_pr <- "<p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr>"
  nv_sp_pr <- sprintf( nv_sp_pr, ifelse(!is.null(location$ph) && !is.na(location$ph), location$ph, "") )
  xml_elt <- paste0( pml_with_ns("p:sp"),
                     nv_sp_pr, sp_pr,
                     "<p:txBody><a:bodyPr/><a:lstStyle/>",
                     pars,
                     "</p:txBody></p:sp>"
  )
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
  ph_with_img( x, src = file, location = location )
}
