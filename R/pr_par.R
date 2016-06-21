#' @title Paragraph formatting properties
#'
#' @description Create a \code{pr_par} object that describes
#' paragraph formatting properties.
#'
#' @details pr_par is used to control paragraph properties.
#' It is used when adding plots or when adding content in a FlexTable.
#'
#' @param text.align text alignment - a single character value, expected value
#' is one of 'left', 'right', 'center', 'justify'.
#' @param padding.bottom,padding.top,padding.left,padding.right paragraph paddings - 0 or positive integer value.
#' @param padding paragraph paddings - 0 or positive integer value. Argument \code{padding} overwrites
#' arguments \code{padding.bottom}, \code{padding.top}, \code{padding.left}, \code{padding.right}.
#' @param border shortcut for all borders.
#' @param border.bottom,border.left,border.top,border.right \code{\link{pr_border}} for
#' borders. overwrite other border properties.
#' @param shading.color shading color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @return a \code{pr_par} object
#' @examples
#' pr_par(text.align = "center", padding = 5)
#' @export
#' @importFrom purrr map
pr_par = function(text.align = "left",
                  padding = 0,
                  border = pr_border(width=0),
                  padding.bottom, padding.top,
                  padding.left, padding.right,
                  border.bottom, border.left,
                  border.top, border.right,
                  shading.color = "transparent") {

	out = list()

	out <- check_set_color(out, shading.color)
	out <- check_set_choice( obj = out, value = text.align,
	                         choices = c("left", "right", "center", "justify") )
	# padding checking
	out <- check_spread_integer( out, padding,
	                             c("padding.bottom", "padding.top",
	                               "padding.left", "padding.right"))
	if( !missing(padding.bottom) )
	  out <- check_set_integer( obj = out, padding.bottom)
	if( !missing(padding.left) )
	  out <- check_set_integer( obj = out, padding.left)
	if( !missing(padding.top) )
	  out <- check_set_integer( obj = out, padding.top)
	if( !missing(padding.right) )
	  out <- check_set_integer( obj = out, padding.right)

	# border checking
	out <- check_spread_border( obj = out, border,
	                            dest = c("border.bottom", "border.top",
	                                     "border.left", "border.right") )
	if( !missing(border.top) )
	  out <- check_set_border( obj = out, border.top)
	if( !missing(border.bottom) )
	  out <- check_set_border( obj = out, border.bottom)
	if( !missing(border.left) )
	  out <- check_set_border( obj = out, border.left)
	if( !missing(border.right) )
	  out <- check_set_border( obj = out, border.right)

	class( out ) = "pr_par"

	out
}

#' @export
#' @importFrom purrr map_chr
#' @importFrom purrr map_int
#' @importFrom grDevices col2rgb
format.pr_par = function (x, type = "wml", ...){
  btlr_list <- list(x$border.bottom, x$border.top,
                    x$border.left, x$border.right)

  btlr_cols <- map( btlr_list,
       function(x) as.vector(col2rgb(x$color, alpha = TRUE )[,1] ) )
  colmat <- do.call( "rbind", btlr_cols )
  types <- map_chr( btlr_list, "style" )
  widths <- map_int( btlr_list, "width" )
  shading <- col2rgb(x$shading.color, alpha = TRUE )[,1]

  stopifnot(length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html") )


  if( type == "wml" ){
    w_ppr(text_align = x$text.align,
        pb = x$padding.bottom, pt = x$padding.top,
        pl = x$padding.left, pr = x$padding.right,
        shd_r = shading[1], shd_g = shading[2], shd_b = shading[3], shd_a = shading[4],
        colmat[,1], colmat[,2], colmat[,3], colmat[,4],
        type = types, width = widths)
  } else if( type == "pml" ){
    a_ppr(text_align = x$text.align,
          pb = x$padding.bottom, pt = x$padding.top,
          pl = x$padding.left, pr = x$padding.right,
          shd_r = shading[1], shd_g = shading[2], shd_b = shading[3], shd_a = shading[4],
          colmat[,1], colmat[,2], colmat[,3], colmat[,4],
          type = types, width = widths)
  } else if( type == "html" ){
    css_ppr(text_align = x$text.align,
          pb = x$padding.bottom, pt = x$padding.top,
          pl = x$padding.left, pr = x$padding.right,
          shd_r = shading[1], shd_g = shading[2], shd_b = shading[3], shd_a = shading[4],
          colmat[,1], colmat[,2], colmat[,3], colmat[,4],
          type = types, width = widths)
  } else stop("unimplemented")

}




#' @param x,object \code{pr_par} object
#' @param ... further arguments - not used
#' @rdname pr_par
#' @export
print.pr_par = function (x, ...){
	cat( format(x, type = "html") )
}


#' @rdname pr_par
#' @examples
#' obj <- pr_par(text.align = "center", padding = 1)
#' update( obj, padding.bottom = 5 )
#' @export
update.pr_par <- function(object, text.align, padding, border,
                          padding.bottom, padding.top, padding.left, padding.right,
                          border.bottom, border.left,border.top, border.right,
                          shading.color, ...) {

  if( !missing( text.align ) )
    object <- check_set_choice( obj = object, value = text.align,
                             choices = c("left", "right", "center", "justify") )

  # padding checking
  if( !missing( padding ) )
    object <- check_spread_integer( object, padding,
               c("padding.bottom", "padding.top",
                 "padding.left", "padding.right"))
  if( !missing(padding.bottom) )
    object <- check_set_integer( obj = object, padding.bottom)
  if( !missing(padding.left) )
    object <- check_set_integer( obj = object, padding.left)
  if( !missing(padding.top) )
    object <- check_set_integer( obj = object, padding.top)
  if( !missing(padding.right) )
    object <- check_set_integer( obj = object, padding.right)

  # border checking
  if( !missing( border ) )
    object <- check_spread_border( obj = object, border,
                              dest = c("border.bottom", "border.top",
                                       "border.left", "border.right") )
  if( !missing(border.top) )
    object <- check_set_border( obj = object, border.top)
  if( !missing(border.bottom) )
    object <- check_set_border( obj = object, border.bottom)
  if( !missing(border.left) )
    object <- check_set_border( obj = object, border.left)
  if( !missing(border.right) )
    object <- check_set_border( obj = object, border.right)

  if( !missing( shading.color ) )
    object <- check_set_color(object, shading.color)

  object
}
