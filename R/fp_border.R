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


#' @param object fp_border object
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
print.fp_border <- function(x, ...) {

  msg <- paste0("line: color: ", x$color, ", width: ", x$width, ", style: ", x$style, "\n")
  cat(msg)
  invisible()
}
