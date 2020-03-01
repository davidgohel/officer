#' @export
#' @title add a new empty shape
#' @description add a new empty shape in the current slide.
#' This function is deprected, function `ph_with` should be used
#' instead.
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
ph_empty <- function( x, type = "body", index = 1, location = NULL ){
  .Deprecated(new = "ph_with")
  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
  }
  ph_with(x, empty_content(), location = location)
}

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
  ph_empty(x, location = location)
}


#' @export
#' @title add text into a new shape
#' @description add text into a new shape in a slide.
#' This function is deprecated in favor of \code{\link{ph_with}}.
#' @inheritParams ph_empty
#' @param str text to add
#' @inherit ph_empty seealso
ph_with_text <- function( x, str, type = "title", index = 1, location = NULL ){
  .Deprecated(new = "ph_with")

  if(!is.character(str))
    str <- format(str)
  if( is.null( location ) ){
    location <- ph_location_type(type = type, id = index)
  }
  ph_with(x, str, location = location)
}



#' @export
#' @title add image
#' @description add an image as a new shape in the current slide.
#' This function is deprecated in favor of \code{\link{ph_with}}.
#' @inheritParams ph_empty
#' @param src image filename, the basename of the file must not contain any blank.
#' @param width,height image size in inches
#' @param left,top location of the new shape on the slide
#' @param rot rotation angle
ph_with_img_at <- function( x, src, left, top, width, height, rot = 0 ){
  .Deprecated(new = "ph_with")
  ph_with(x, external_img(src=src, width = width, height = height),
          location = ph_location(ph = "", label = "", left = left, top = top, width = width, height = height, rotation = rot))

}


#' @export
#' @title add ggplot to a pptx presentation
#' @description add a ggplot as a png image into an rpptx object
#' This function is deprecated in favor of \code{\link{ph_with}}.
#' @inheritParams ph_empty
#' @param value ggplot object
#' @param width,height image size in inches
#' @param left,top location of the new shape on the slide
#' @param ... Arguments to be passed to png function.
#' @importFrom grDevices png dev.off
ph_with_gg_at <- function( x, value, width, height, left, top, ... ){

  .Deprecated(new = "ph_with")
  if( !requireNamespace("ggplot2") )
    stop("package ggplot2 is required to use this function")

  stopifnot(inherits(value, "gg"))

  ph_with(x,
          value,
          location = ph_location(
            ph = "", label = "", left = left, top = top,
            width = width, height = height), ...)
}






