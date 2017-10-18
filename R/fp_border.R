border_styles = c( "none", "solid", "dotted", "dashed" )

#' @title border properties object
#'
#' @description create a border properties object.
#'
#' @param color border color - single character value (e.g. "#000000" or "black")
#' @param style border style - single character value : "none" or "solid" or "dotted" or "dashed"
#' @param width border width - an integer value : 0>= value
#' @examples
#' fp_border()
#' fp_border(color="orange", style="solid", width=1)
#' fp_border(color="gray", style="dotted", width=1)
#' @export
fp_border = function( color = "black", style = "solid", width = 1 ){

  out <- list()
  out <- check_set_numeric( obj = out, width)
  out <- check_set_color(out, color)
  out <- check_set_choice( obj = out, style,
                           choices = border_styles )

	class( out ) = "fp_border"
	out
}


#' @param ... further arguments - not used
#' @rdname fp_border
#' @examples
#'
#' # modify object ------
#' border <- fp_border()
#' update(border, style="dotted", width=3)
#' @export
update.fp_border <- function(object, color, style, width, ...) {


  if( !missing( color ) ){
    object <- check_set_color(object, color)
  }

  if( !missing( width ) ){
    object <- check_set_integer( obj = object, width)
  }

  if( !missing( style ) ){
    object <- check_set_choice( obj = object, style, choices = border_styles )
  }

  object
}


#' @export
#' @rdname fp_border
#' @param x,object \code{fp_border} object
#' @param type output type - only 'pml' currently supported.
format.fp_border = function (x, type = "pml", ...){


  col_mat <- col2rgb(x$color, alpha = TRUE)
  red <- col_mat[1,1]
  green <- col_mat[2,1]
  blue <- col_mat[3,1]
  alpha <- col_mat[4,1]


  if( type == "pml"){

    a_border(r = red, g = green, b = blue, a = alpha, type = x$style, width = x$width)

  } else stop("unimplemented")
}
