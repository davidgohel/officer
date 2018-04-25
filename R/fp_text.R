#' @title Text formatting properties
#'
#' @description Create a \code{fp_text} object that describes
#' text formatting properties.
#'
#' @param color font color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @param font.size font size (in point) - 0 or positive integer value.
#' @param bold is bold
#' @param italic is italic
#' @param underlined is underlined
#' @param font.family single character value specifying font name.
#' @param vertical.align single character value specifying font vertical alignments.
#' Expected value is one of the following : default \code{'baseline'}
#' or \code{'subscript'} or \code{'superscript'}
#' @param shading.color shading color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @return a \code{fp_text} object
#' @export
fp_text <- function( color = "black", font.size = 10,
                    bold = FALSE, italic = FALSE, underlined = FALSE,
                    font.family = "Arial",
                    vertical.align = "baseline",
                    shading.color = "transparent" ){

  out <- list()

  out <- check_set_integer( obj = out, font.size)
  out <- check_set_bool( obj = out, bold)
  out <- check_set_bool( obj = out, italic)
  out <- check_set_bool( obj = out, underlined)
  out <- check_set_color(out, color)
  out <- check_set_chr(out, font.family)

  out <- check_set_choice( obj = out, value = vertical.align,
                           choices = c("subscript", "superscript", "baseline") )
  out <- check_set_color(out, shading.color)

  class( out ) <- "fp_text"

  out
}

rpr_pointer <- function(x){
  cols <- as.integer( col2rgb(x$color, alpha = TRUE)[,1] )
  shadings <- as.integer( col2rgb(x$shading.color, alpha = TRUE)[,1] )

  x <- append(x, list(
    col_font_r = cols[1], col_font_g = cols[2],
    col_font_b = cols[3], col_font_a = cols[4],
    col_shading_r = shadings[1], col_shading_g = shadings[2],
    col_shading_b = shadings[3], col_shading_a = shadings[4]
  ))
  rpr_new( x )
}

#' @rdname fp_text
#' @param format format type, wml for MS word, pml for
#' MS PowerPoint and html.
#' @param type output type - one of 'wml', 'pml', 'html'.
#' @export
format.fp_text <- function( x, type = "wml", ... ){

  stopifnot(length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html") )

  ptr <- rpr_pointer( x )

  if( type == "wml" ){
    rpr_w(ptr)
  } else if( type == "pml" ){
    rpr_p(ptr)
  } else if(type == "html") {
    rpr_css(ptr)
  } else stop("unimplemented type")
}

#' @param x \code{fp_text} object
#' @examples
#' print( fp_text (color="red", font.size = 12) )
#' @rdname fp_text
#' @export
print.fp_text = function (x, ...){
  out <- data.frame(
    size = as.double(x$font.size),
    italic = x$italic,
    bold = x$bold,
    underlined = x$underlined,
    color = x$color,
    shading = x$shading.color,
    fontname = x$font.family,
    vertical_align = x$vertical.align, stringsAsFactors = FALSE )
  print(out)
  invisible()
}


#' @param object \code{fp_text} object to modify
#' @param ... further arguments - not used
#' @rdname fp_text
#' @export
update.fp_text <- function(object, color, font.size,
                           bold = FALSE, italic = FALSE, underlined = FALSE,
                           font.family, vertical.align, shading.color, ...) {

  if( !missing( font.size ) )
    object <- check_set_integer( obj = object, font.size)
  if( !missing( bold) )
    object <- check_set_bool( obj = object, bold)
  if( !missing( italic) )
    object <- check_set_bool( obj = object, italic)
  if( !missing( underlined) )
    object <- check_set_bool( obj = object, underlined)
  if( !missing( color ) )
    object <- check_set_color(object, color)
  if( !missing( font.family ) )
    object <- check_set_chr(object, font.family)
  if( !missing( vertical.align ) )
    object <- check_set_choice(
      obj = object, value = vertical.align,
      choices = c("subscript", "superscript", "baseline") )
  if( !missing(shading.color) )
    object <- check_set_color(object, shading.color)

  object
}

