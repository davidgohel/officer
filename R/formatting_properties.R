# check properties helpers ----

check_spread_integer <- function( obj, value, dest){
  varname <- as.character(substitute(value))
  if( is.numeric( value ) && length(value) == 1  && value >= 0 ){
    for(i in dest)
      obj[[i]] <- as.integer(value)
  } else stop(varname, " must be a positive integer scalar.", call. = FALSE)
  obj
}

check_set_color <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !is.color( value ) && !is.na(value) )
    stop(varname, " must be a valid color.", call. = FALSE )
  else obj[[varname]] <- value
  obj
}

check_set_pic <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !grepl(pattern = "^rId[0-9]+", value) )
    stop(varname, " must be a valid reference id: ", value, call. = FALSE )
  obj[[varname]] <- value
  obj
}
check_set_file <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !file.exists(value) )
    stop(varname, " must be a valid filename.", call. = FALSE )
  obj[[varname]] <- value
  obj
}
check_set_border <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !inherits( value, "fp_border" ) )
    stop(varname, " must be a fp_border object." , call. = FALSE)
  else obj[[varname]] <- value
  obj
}

check_spread_border <- function( obj, value, dest ){
  varname <- as.character(substitute(value))
  if( !inherits( value, "fp_border" ) )
    stop(varname, " must be a fp_border object." , call. = FALSE)
  for(i in dest )
    obj[[i]] <- value
  obj
}

check_set_integer <- function( obj, value){
  varname <- as.character(substitute(value))
  if( is.na(value) || (is.numeric( value ) && length(value) == 1  && value >= 0) ){
    obj[[varname]] <- as.integer(value)
  } else stop(varname, " must be a positive integer scalar.", call. = FALSE)
  obj
}

check_set_numeric <- function( obj, value){
  varname <- as.character(substitute(value))
  if( is.na(value) || (is.numeric( value ) && length(value) == 1  && value >= 0) ){
    obj[[varname]] <- as.double(value)
  } else stop(varname, " must be a positive numeric scalar.", call. = FALSE)
  obj
}

check_set_bool <- function( obj, value){
  varname <- as.character(substitute(value))
  if( is.na(value) || (is.logical( value ) && length(value) == 1) ){
    obj[[varname]] <- value
  } else stop(varname, " must be a boolean", call. = FALSE)
  obj
}
check_set_chr <- function( obj, value){
  varname <- as.character(substitute(value))
  if( is.na(value) || (is.character( value ) && length(value) == 1) ){
    obj[[varname]] <- value
  } else stop(varname, " must be a string", call. = FALSE)
  obj
}

check_set_choice <- function( obj, value, choices){
  varname <- as.character(substitute(value))

  if(is.na(value)){
    obj[[varname]] <- value
  } else {
    if( is.character( value ) && length(value) == 1 ){
      if( !value %in% choices )
        stop(varname, " must be one of ",
             paste( shQuote(choices), collapse = ", "),
             call. = FALSE )
      obj[[varname]] = value
    } else stop(varname, " must be a character scalar.", call. = FALSE)
  }

  obj
}

check_set_class <- function( obj, value, cl){
  varname <- as.character(substitute(value))
  cl_str <- sprintf(" must be a %s object.", cl)
  if( !inherits( value, cl ) )
    stop(varname, cl_str, call. = FALSE)
  else obj[[varname]] <- value
  obj
}

# fp_text ----
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
#' @param font.family single character value. Specifies the font to
#' be used to format characters in the Unicode range (U+0000-U+007F).
#' @param cs.family optional font to be used to format
#' characters in a complex script Unicode range. For example, Arabic
#' text might be displayed using the "Arial Unicode MS" font.
#' @param eastasia.family optional font to be used to
#' format characters in an East Asian Unicode range. For example,
#' Japanese text might be displayed using the "MS Mincho" font.
#' @param hansi.family optional. Specifies the font to be used to format
#' characters in a Unicode range which does not fall into one of the
#' other categories.
#' @param vertical.align single character value specifying font vertical alignments.
#' Expected value is one of the following : default \code{'baseline'}
#' or \code{'subscript'} or \code{'superscript'}
#' @param shading.color shading color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @return a \code{fp_text} object
#' @examples
#' fp_text()
#' fp_text(color = "red")
#' fp_text(bold = TRUE, shading.color = "yellow")
#' @family functions for defining formatting properties
#' @seealso [ftext], [fpar]
#' @export
fp_text <- function(color = "black", font.size = 10,
                    bold = FALSE, italic = FALSE, underlined = FALSE,
                    font.family = "Arial",
                    cs.family = NULL, eastasia.family = NULL, hansi.family = NULL,
                    vertical.align = "baseline",
                    shading.color = "transparent" ){

  out <- list()

  out <- check_set_numeric( obj = out, font.size)
  out <- check_set_bool( obj = out, bold)
  out <- check_set_bool( obj = out, italic)
  out <- check_set_bool( obj = out, underlined)
  out <- check_set_color(out, color)
  out <- check_set_chr(out, font.family)

  if(is.null(cs.family)) cs.family <- font.family
  if(is.null(eastasia.family)) eastasia.family <- font.family
  if(is.null(hansi.family)) hansi.family <- font.family

  out <- check_set_chr(out, cs.family)
  out <- check_set_chr(out, eastasia.family)
  out <- check_set_chr(out, hansi.family)

  out <- check_set_choice( obj = out, value = vertical.align,
                           choices = c("subscript", "superscript", "baseline") )
  out <- check_set_color(out, shading.color)

  class( out ) <- "fp_text"

  out
}

#' @rdname fp_text
#' @description Function `fp_text_lite()` is generating properties
#' with only entries for the parameters users provided. The
#' undefined properties will inherit from the default settings.
#' @export
fp_text_lite <- function(
  color = NA, font.size = NA,
  font.family = NA, cs.family = NA, eastasia.family = NA, hansi.family = NA,
  bold = NA, italic = NA, underlined = NA,
  vertical.align = "baseline", shading.color = NA){

  fp_text(
    color = color, font.size = font.size,
    bold = bold, italic = italic, underlined = underlined,
    font.family = font.family, cs.family = cs.family, eastasia.family = eastasia.family, hansi.family = hansi.family,
    vertical.align = vertical.align, shading.color = shading.color)

}

#' @rdname fp_text
#' @param format format type, wml for MS word, pml for
#' MS PowerPoint and html.
#' @param type output type - one of 'wml', 'pml', 'html'.
#' @export
format.fp_text <- function( x, type = "wml", ... ){

  stopifnot(length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html") )

  if( type == "wml" ){
    rpr_wml(x)
  } else if( type == "pml" ){
    rpr_pml(x)
  } else if(type == "html") {
    rpr_css(x)
  } else stop("unimplemented type")
}

#' @export
to_wml.fp_text <- function(x, add_ns = FALSE, ...) {
  format(x, type = "wml")
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
    fontname_cs = x$cs.family,
    fontname_eastasia = x$eastasia.family,
    fontname.hansi = x$hansi.family,
    vertical_align = x$vertical.align, stringsAsFactors = FALSE )
  print(out)
  invisible()
}


#' @param object \code{fp_text} object to modify
#' @param ... further arguments - not used
#' @rdname fp_text
#' @export
update.fp_text <- function(object, color, font.size,
                           bold, italic, underlined,
                           font.family, cs.family, eastasia.family, hansi.family,
                           vertical.align, shading.color, ...) {

  if( !missing( font.size ) )
    object <- check_set_numeric( obj = object, font.size)
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
  if( !missing( cs.family ) )
    object <- check_set_chr(object, cs.family)
  if( !missing( eastasia.family ) )
    object <- check_set_chr(object, eastasia.family)
  if( !missing( hansi.family ) )
    object <- check_set_chr(object, hansi.family)
  if( !missing( vertical.align ) )
    object <- check_set_choice(
      obj = object, value = vertical.align,
      choices = c("subscript", "superscript", "baseline") )
  if( !missing(shading.color) )
    object <- check_set_color(object, shading.color)

  object
}

# fp_border ----
border_styles = c( "none", "solid", "dotted", "dashed" )

#' @title Border properties object
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
#' @family functions for defining formatting properties
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

# fp_par -----
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
#' @param line_spacing line spacing, 1 is single line spacing, 2 is double line spacing.
#' @param border shortcut for all borders.
#' @param border.bottom,border.left,border.top,border.right \code{\link{fp_border}} for
#' borders. overwrite other border properties.
#' @param shading.color shading color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @param keep_with_next a scalar logical. Specifies that the paragraph (or at least part of it) should be rendered
#' on the same page as the next paragraph when possible.
#' @return a \code{fp_par} object
#' @examples
#' fp_par(text.align = "center", padding = 5)
#' @export
#' @family functions for defining formatting properties
#' @seealso [fpar]
fp_par = function(text.align = "left",
                  padding = 0,
                  line_spacing = 1,
                  border = fp_border(width=0),
                  padding.bottom, padding.top,
                  padding.left, padding.right,
                  border.bottom, border.left,
                  border.top, border.right,
                  shading.color = "transparent",
                  keep_with_next = FALSE) {

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

  out <- check_set_numeric( obj = out, line_spacing)

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

  out$keep_with_next <- keep_with_next
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
#' @export
to_wml.fp_par <- function(x, add_ns = FALSE, ...) {
  format(x, type = "wml")
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

# fp_cell ----

vertical.align.styles <- c( "top", "center", "bottom" )
text.directions <- c( "lrtb", "tbrl", "btlr" )

#' @title Cell formatting properties
#'
#' @description Create a \code{fp_cell} object that describes cell formatting properties.
#'
#' @param border shortcut for all borders.
#' @param border.bottom,border.left,border.top,border.right \code{\link{fp_border}} for borders.
#' @param vertical.align cell content vertical alignment - a single character value,
#' expected value is one of "center" or "top" or "bottom"
#' @param margin shortcut for all margins.
#' @param margin.bottom,margin.top,margin.left,margin.right cell margins - 0 or positive integer value.
#' @param background.color cell background color - a single character value specifying a
#' valid color (e.g. "#000000" or "black").
#' @param text.direction cell text rotation - a single character value, expected
#' value is one of "lrtb", "tbrl", "btlr".
#' @export
#' @family functions for defining formatting properties
fp_cell = function(
  border = fp_border(width=0),
  border.bottom, border.left, border.top, border.right,
  vertical.align = "center",
  margin = 0,
  margin.bottom, margin.top, margin.left, margin.right,
  background.color = "transparent",
  text.direction = "lrtb"){

  out <- list()

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

  # background-color checking
  out <- check_set_color(out, background.color)

  out <- check_set_choice( obj = out, value = vertical.align,
                           choices = vertical.align.styles )
  out <- check_set_choice( obj = out, value = text.direction,
                           choices = text.directions )

  # margin checking
  out <- check_spread_integer( out, margin,
                               c("margin.bottom", "margin.top",
                                 "margin.left", "margin.right"))

  if( !missing(margin.bottom) )
    out <- check_set_integer( obj = out, margin.bottom)
  if( !missing(margin.left) )
    out <- check_set_integer( obj = out, margin.left)
  if( !missing(margin.top) )
    out <- check_set_integer( obj = out, margin.top)
  if( !missing(margin.right) )
    out <- check_set_integer( obj = out, margin.right)


  class( out ) = "fp_cell"
  out
}


#' @export
#' @rdname fp_cell
#' @param x,object \code{fp_cell} object
#' @param type output type - one of 'wml', 'pml', 'html'.
#' @param ... further arguments - not used
format.fp_cell = function (x, type = "wml", ...){
  btlr_list <- list(x$border.bottom, x$border.top,
                    x$border.left, x$border.right)

  btlr_cols <- lapply( btlr_list,
                       function(x) {
                         as.vector(col2rgb(x$color, alpha = TRUE )[,1] )
                       }
  )
  colmat <- do.call( "rbind", btlr_cols )
  types <- sapply(btlr_list, function(x) x$style )
  widths <- sapply(btlr_list, function(x) x$width )
  shading <- col2rgb(x$background.color, alpha = TRUE )[,1]

  if( type == "wml"){
    tcpr_wml(x)
  } else if( type == "pml"){
    tcpr_pml(x)
  } else if( type == "html" ){
    tcpr_css(x)
  } else stop("unimplemented")
}

#' @export
to_wml.fp_cell <- function(x, add_ns = FALSE, ...) {
  format(x, type = "wml")
}

#' @export
#' @rdname fp_cell
print.fp_cell <- function(x, ...){
  cat(format(x, type = "html"))
}



#' @rdname fp_cell
#' @examples
#' obj <- fp_cell(margin = 1)
#' update( obj, margin.bottom = 5 )
#' @export
update.fp_cell <- function(object, border,
                           border.bottom,border.left,border.top,border.right,
                           vertical.align, margin = 0,
                           margin.bottom, margin.top, margin.left, margin.right,
                           background.color,
                           text.direction, ...) {

  if( !missing(border) )
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

  # background-color checking
  if( !missing(background.color) )
    object <- check_set_color(object, background.color)

  if( !missing(vertical.align) )
    object <- check_set_choice( obj = object, value = vertical.align,
                                choices = vertical.align.styles )
  if( !missing(text.direction) )
    object <- check_set_choice( obj = object, value = text.direction,
                                choices = text.directions )

  # margin checking
  if( !missing(margin) )
    object <- check_spread_integer( object, margin,
                                    c("margin.bottom", "margin.top",
                                      "margin.left", "margin.right"))

  if( !missing(margin.bottom) )
    object <- check_set_integer( obj = object, margin.bottom)
  if( !missing(margin.left) )
    object <- check_set_integer( obj = object, margin.left)
  if( !missing(margin.top) )
    object <- check_set_integer( obj = object, margin.top)
  if( !missing(margin.right) )
    object <- check_set_integer( obj = object, margin.right)

  object
}
