# utils ----


css_color <- function(color){
  color <- as.vector(col2rgb(color, alpha = TRUE)) / c(1, 1, 1, 255)

  if( !(color[4] > 0) ) "transparent"
  else sprintf("rgba(%.0f,%.0f,%.0f,%.2f)",
               color[1], color[2], color[3], color[4])
}
hex_color <- function(color){
  color <- as.vector(col2rgb(color, alpha = TRUE)) / c(1, 1, 1, 255)
  sprintf("%02X%02X%02X",
          color[1], color[2], color[3])
}
is_transparent <- function(color){
  if("transparent" %in% color) return(TRUE)
  color <- as.vector(col2rgb(color, alpha = TRUE))[4] / 255
  !(color > 0)
}

border_wml <- function(x, side){
  tagname <- paste0("w:", side)
  if( !x$style %in% c("dotted", "dashed", "solid") ){
    x$style <- "solid"
  }
  style_ <- sprintf("w:val=\"%s\"", x$style)

  width_ <- sprintf("w:sz=\"%.0f\"", x$width*8)
  color_ <- sprintf("w:color=\"%s\"", hex_color(x$color))

  paste0("<", tagname,
         " ", style_,
         " ", width_,
         " ", "w:space=\"0\"",
         " ", color_,
         "/>")

}
border_css <- function(x, side){

  color_ <- css_color(x$color)
  if( !(x$width > 0 ) )
    color_ <- "transparent"

  width_ <- sprintf("%.02fpt", x$width)

  if( !x$style %in% c("dotted", "dashed", "solid") ){
    x$style <- "solid"
  }
  paste0("border-", side, ": ", width_, " ", x$style, " ", color_)

}

ppr_pml <- function(x){
  align  <- " algn=\"r\""
  if (x$text.align == "left" ){
    align  <- " algn=\"l\"";
  } else if(x$text.align == "center" ){
    align  <- " algn=\"ctr\"";
  } else if(x$text.align == "justify" ){
    align  <- " algn=\"just\"";
  }
  leftright_padding <- sprintf(" marL=\"%.0f\" marR=\"%.0f\"", x$padding.left*12700, x$padding.right*12700)
  top_padding <- sprintf("<a:spcBef><a:spcPts val=\"%.0f\" /></a:spcBef>", x$padding.top*100)
  bottom_padding <- sprintf("<a:spcAft><a:spcPts val=\"%.0f\" /></a:spcAft>", x$padding.bottom*100)

  paste0("<a:pPr", align, leftright_padding, ">",
         top_padding, bottom_padding, "<a:buNone/>",
         "</a:pPr>")
}

ppr_css <- function(x){

  text.align  <- sprintf("text-align:%s;", x$text.align)
  borders <- paste0(
    border_css(x$border.bottom, "bottom"),
    border_css(x$border.top, "top"),
    border_css(x$border.left, "left"),
    border_css(x$border.right, "right") )

  paddings <- sprintf("padding-top:%.0fpt;padding-bottom:%.0fpt;padding-left:%.0fpt;padding-right:%.0fpt;",
          x$padding.top, x$padding.bottom, x$padding.left, x$padding.right)

  shading.color <- sprintf("background-color:%s;", css_color(x$shading.color))

  paste0("margin:0pt;", text.align, borders, paddings,
         shading.color)
}

ppr_wml <- function(x){

  if("justify" %in% x$text.align ){
    x$text.align  <- "both";
  }
  text_align_ <- sprintf("<w:jc w:val=\"%s\"/>", x$text.align)
  borders_ <- paste0(
    border_wml(x$border.bottom, "bottom"),
    border_wml(x$border.top, "top"),
    border_wml(x$border.left, "left"),
    border_wml(x$border.right, "right") )

  leftright_padding <- sprintf("<w:ind w:firstLine=\"0\" w:left=\"%.0f\" w:right=\"%.0f\"/>",
                               x$padding.left*20, x$padding.right*20)
  topbot_spacing <- sprintf("<w:spacing w:after=\"%.0f\" w:before=\"%.0f\"/>",
                            x$padding.bottom*20, x$padding.top*20)
  shading_ <- ""
  if(!is_transparent(x$shading.color)){
    shading_ <- sprintf(
      "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"%s\"/>",
      hex_color(x$shading.color))
  }

  paste0("<w:pPr>",
         text_align_,
         borders_,
         topbot_spacing,
         leftright_padding,
         shading_,
         "</w:pPr>")

}







# main ----


#' @title Paragraph formatting properties
#'
#' @description Create a \code{fp_par} object that describes
#' paragraph formatting properties.
#'
#' @param text.align text alignment - a single character value, expected value
#' is one of 'left', 'right', 'center', 'justify'.
#' @param padding.bottom,padding.top,padding.left,padding.right paragraph paddings - 0 or positive integer value.
#' @param padding paragraph paddings - 0 or positive integer value. Argument \code{padding} overwrites
#' arguments \code{padding.bottom}, \code{padding.top}, \code{padding.left}, \code{padding.right}.
#' @param border shortcut for all borders.
#' @param border.bottom,border.left,border.top,border.right \code{\link{fp_border}} for
#' borders. overwrite other border properties.
#' @param shading.color shading color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @return a \code{fp_par} object
#' @examples
#' fp_par(text.align = "center", padding = 5)
#' @export
fp_par = function(text.align = "left",
                  padding = 0,
                  border = fp_border(width=0),
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

	class( out ) = "fp_par"

	out
}

#' @export
#' @importFrom grDevices col2rgb
format.fp_par = function (x, type = "wml", ...){

  stopifnot(length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html") )

  if( type == "wml" ){
    ppr_wml(x)
  } else if( type == "pml" ){
    ppr_pml(x)
  } else if( type == "html" ){
    ppr_css(x)
  } else stop("unimplemented")

}


#' @rdname fp_par
#' @export
#' @importFrom stats setNames
dim.fp_par = function (x){
  width <- x$padding.left + x$padding.right #+ x$border.left$width + x$border.right$width
  height <- x$padding.top + x$padding.bottom #+ x$border.top$width + x$border.bottom$width
  setNames(c(width, height) * (4/3) / 72, c("width", "height"))
}



#' @param x,object \code{fp_par} object
#' @param ... further arguments - not used
#' @rdname fp_par
#' @export
print.fp_par = function (x, ...){
  out <- data.frame(
    text.align = as.character(x$text.align),
    padding.top = as.character(x$padding.top),
    padding.bottom = as.character(x$padding.bottom),
    padding.left = as.character(x$padding.left),
    padding.right = as.character(x$padding.right),
    shading.color = as.character(x$shading.color) )
  out <- as.data.frame( t(out) )
  names(out) <- "values"
  print(out)
  cat("borders:\n")
  borders <- rbind(
  as.data.frame( unclass(x$border.top )),
  as.data.frame( unclass(x$border.bottom )),
  as.data.frame( unclass(x$border.left )),
  as.data.frame( unclass(x$border.right )) )
  row.names(borders) = c("top", "bottom", "left", "right")
  print(borders)
}


#' @rdname fp_par
#' @examples
#' obj <- fp_par(text.align = "center", padding = 1)
#' update( obj, padding.bottom = 5 )
#' @export
update.fp_par <- function(object, text.align, padding, border,
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
