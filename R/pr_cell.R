vertical.align.styles <- c( "top", "middle", "bottom" )
text.directions <- c( "lrtb", "tbrl", "btlr" )

#' @importFrom gdtools raster_str

#' @title Cell formatting properties
#'
#' @description Create a \code{pr_cell} object that describes cell formatting properties.
#'
#' @param border shortcut for all borders.
#' @param border.bottom,border.left,border.top,border.right \code{\link{pr_border}} for borders.
#' @param vertical.align cell content vertical alignment - a single character value
#' , expected value is one of "center" or "top" or "bottom"
#' @param margin shortcut for all margins.
#' @param margin.bottom,margin.top,margin.left,margin.right cell margins - 0 or positive integer value.
#' @param background.color cell background color - a single character value specifying a
#' valid color (e.g. "#000000" or "black").
#' @param text.direction cell text rotation - a single character value, expected
#' value is one of "lrtb", "tbrl", "btlr".
#' @param row_span,column_span row and column span
#' @export
pr_cell = function(
  border = pr_border(width=0),
	border.bottom,border.left,border.top,border.right,
	vertical.align = "middle",
  margin = 0,
	margin.bottom, margin.top, margin.left, margin.right,
  background.color = "transparent",
  background.img.id = "rId1",
  background.img.src = NULL,
  text.direction = "lrtb",
	row_span = 1L, column_span = 1L
){

out <- list()

out <- check_set_integer( obj = out, row_span)
out <- check_set_integer( obj = out, column_span)

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

if( !is.null( background.img.src ) ){
  out <- check_set_pic(out, background.img.id)
  out <- check_set_file(out, background.img.src)
}

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


class( out ) = "pr_cell"
out
}


#' @export
#' @rdname pr_cell
#' @param x,object object \code{pr_cell}
#' @param type output type - one of 'wml', 'pml', 'html'.
#' @param ... further arguments - not used
format.pr_cell = function (x, type = "wml", ...){
  btlr_list <- list(x$border.bottom, x$border.top,
                    x$border.left, x$border.right)

  btlr_cols <- map( btlr_list,
                    function(x) {
                      as.vector(col2rgb(x$color, alpha = TRUE )[,1] )
                    }
  )
  colmat <- do.call( "rbind", btlr_cols )
  types <- map_chr( btlr_list, "style" )
  widths <- map_int( btlr_list, "width" )
  shading <- col2rgb(x$background.color, alpha = TRUE )[,1]

  if( !is.null( x$background.img.id )){
    do_bg_img <- TRUE
    img_id <- x$background.img.id
    img_src <- x$background.img.src
  } else {
    do_bg_img <- FALSE
    img_id <- ""
    img_src <- ""
  }


  if( type == "wml"){

    w_tcpr(vertical_align = x$vertical.align,
      text_direction = x$text.direction,
      mb = x$margin.bottom, mt = x$margin.top,
      ml = x$margin.left, mr = x$margin.right,
      shd_r = shading[1], shd_g = shading[2], shd_b = shading[3], shd_a = shading[4],
      do_bg_img, img_id, img_src,
      colmat[,1], colmat[,2], colmat[,3], colmat[,4],
      type = types, width = widths,
      row_span = x$row_span, column_span = x$column_span )

  } else if( type == "pml"){

    a_tcpr(vertical_align = x$vertical.align,
           text_direction = x$text.direction,
           mb = x$margin.bottom, mt = x$margin.top,
           ml = x$margin.left, mr = x$margin.right,
           shd_r = shading[1], shd_g = shading[2], shd_b = shading[3], shd_a = shading[4],
           do_bg_img, img_id, img_src,
           colmat[,1], colmat[,2], colmat[,3], colmat[,4],
           type = types, width = widths,
           row_span = x$row_span, column_span = x$column_span )

  } else if( type == "html" ){

    css_tcpr(vertical_align = x$vertical.align,
           text_direction = x$text.direction,
           mb = x$margin.bottom, mt = x$margin.top,
           ml = x$margin.left, mr = x$margin.right,
           shd_r = shading[1], shd_g = shading[2], shd_b = shading[3], shd_a = shading[4],
           do_bg_img, img_id, img_src,
           colmat[,1], colmat[,2], colmat[,3], colmat[,4],
           type = types, width = widths,
           row_span = x$row_span, column_span = x$column_span )
  } else stop("unimplemented")
}

#' @export
#' @rdname pr_cell
print.pr_cell <- function(x, ...){
  cat(format(x, type = "html"))
}



#' @rdname pr_cell
#' @examples
#' obj <- pr_cell(margin = 1)
#' update( obj, margin.bottom = 5 )
#' @export
update.pr_cell <- function(object, border,
                           border.bottom,border.left,border.top,border.right,
                           vertical.align, margin = 0,
                           margin.bottom, margin.top, margin.left, margin.right,
                           background.color, background.img.id, background.img.src,
                           text.direction,
                           row_span, column_span, ...) {

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
  # background-img checking
  if( !missing(background.img.id) )
    object <- check_set_pic(object, background.img.id)
  if( !missing(background.img.src) )
    object <- check_set_file(object, background.img.src)

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

  if( !missing(row_span) )
    object <- check_set_integer( obj = object, row_span)
  if( !missing(column_span) )
    object <- check_set_integer( obj = object, column_span)

  object
}

