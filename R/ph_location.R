# utils ----
get_doc <- function(...){
  args <- list(...)
  x <- args$x

  if( is.null(args$x) ){
    stop("you should not call this function outside of a ph_add call")
  }
  x
}

get_ph_loc <- function(x, layout, master, type, position_right, position_top){
  props <- layout_properties( x, layout = layout, master = master )
  props <- props[props$type %in% type, , drop = FALSE]

  if( nrow(props) < 1) {
    stop("no selected row")
  }

  if(position_right){
    props <- props[props$offx + 0.0001 > max(props$offx),]
  } else {
    props <- props[props$offx - 0.0001 < min(props$offx),]
  }
  if(position_top){
    props <- props[props$offy - 0.0001 < min(props$offy),]
  } else {
    props <- props[props$offy + 0.0001 > max(props$offy),]
  }

  if( nrow(props) > 1) {
    warning("more than a row have been selected")
  }

  props <- props[, c("offx", "offy", "cx", "cy", "ph_label", "ph", "type")]
  names(props) <- c("left", "top", "width", "height", "ph_label", "ph", "type")
  as_ph_location(props)
}

as_ph_location <- function(x, ...){
  if( !is.data.frame(x) ){
    stop("x should be a data.frame")
  }
  ref_names <- c( "width", "height", "left", "top",
                  "ph_label", "ph", "type")
  if (!all(is.element(ref_names, names(x) ))) {
    stop("missing column values:", paste0(setdiff(ref_names, names(x)), collapse = ","))
  }

  out <- x[ref_names]
  as.list(out)
}


# main ----

#' @export
#' @title create a location for a placeholder
#' @description The function will return a list that complies with
#' expected format for argument \code{location} of functions \code{ph_with_*}
#' and \code{ph_with} methods.
#' @param left,top,width,height place holder coordinates
#' in inches.
#' @param label a label for the placeholder. See section details.
#' @param ph openxml string value for tag. This is not meant to be used
#' by end users (unless they understand openxml placeholder and
#' \code{p:ph} tag).
#' @param bg background color
#' @param rotation rotation angle
#' @param ... unused arguments
#' @family functions for placeholder location
#' @details
#' The location of the bounding box associated to a placeholder
#' within a presentation slide is specified with the left top coordinate,
#' the width and the height. These are defined in inches:
#'
#' \describe{
#'   \item{left}{left coordinate of the bounding box}
#'   \item{top}{top coordinate of the bounding box}
#'   \item{width}{width of the bounding box}
#'   \item{height}{height of the bounding box}
#' }
#'
#' In addition to these attributes, there is attribute \code{ph_label}
#' associated with the shape (shapes, text boxes, images and other objects
#' will be identified with that label in the Selection Pane of PowerPoint).
#' This label can then be reused by other functions such as \code{ph_add_fpar},
#' or \code{ph_add_text}.
#'
#' @examples
#' library(magrittr)
#' read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("Hello world",
#'     location = ph_location(width = 4, height = 3, label = "hello") ) %>%
#'   print(target = tempfile(fileext = ".pptx") )
ph_location <- function(left = 1, top = 1, width = 4, height = 3,
                        label = "", ph = "<p:ph/>",
                        bg = NULL, rotation = NULL, ...){

  x <- list(
    left = left,
    top = top,
    width = width,
    height = height,
    ph_label = label,
    ph = ph, bg = bg, rotation = rotation
  )
  x
}

#' @export
#' @title location of a placeholder type
#' @description The function will use the type name of the placeholder (e.g. body, title),
#' the layout name and few other criterias to find the corresponding location.
#' @param type placeholder type to look for in the slide layout, one
#' of 'body', 'title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum'.
#' @param position_right the parameter is used when a selection with above
#' parameters does not provide a unique position (for example
#' layout 'Two Content' contains two element of type 'body').
#' If \code{TRUE}, the element the most on the right side will be selected,
#' otherwise the element the most on the left side will be selected.
#' @param position_top same than \code{position_right} but applied
#' to top versus bottom.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' library(magrittr)
#' read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("Hello world", location = ph_location_type(type = "body") ) %>%
#'   print(target = tempfile(fileext = ".pptx") )
ph_location_type <- function( type = "body", position_right = TRUE, position_top = TRUE, ...){

  ph_types <- c("ctrTitle", "subTitle", "dt", "ftr",
                "sldNum", "title", "body")
  if(!type %in% ph_types){
    stop("argument type must be a value of ", paste0(shQuote(ph_types), collapse = ", ", "."))
  }
  x <- get_doc(...)
  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm()

  args <- list(...)
  layout <- ifelse(is.null(args$layout), unique( xfrm$name ), args$layout)
  master <- ifelse(is.null(args$master), unique( xfrm$master_name ), args$master)
  get_ph_loc(x, layout = layout, master = master,
             type = type, position_right = position_right,
             position_top = position_top)
}

#' @export
#' @title location of a named placeholder
#' @description The function will use the label of a placeholder
#' to find the corresponding location.
#' @param ph_label placeholder label. It can be read in PowerPoint or
#' with function \code{layout_properties()} in column \code{ph_label}.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' library(magrittr)
#' read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("Hello world",
#'     location = ph_location_label(ph_label = "Content Placeholder 2") ) %>%
#'   print(target = tempfile(fileext = ".pptx") )
ph_location_label <- function( ph_label, ...){

  x <- get_doc(...)
  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm()

  layout <- unique( xfrm$name )
  master <- unique(xfrm$master_name)

  props <- layout_properties( x, layout = layout, master = master )
  props <- props[props$ph_label %in% ph_label, , drop = FALSE]

  if( nrow(props) < 1) {
    stop("no selected row")
  }

  if( nrow(props) > 1) {
    warning("more than a row have been selected")
  }

  props <- props[, c("offx", "offy", "cx", "cy", "ph_label", "ph", "type")]
  names(props) <- c("left", "top", "width", "height", "ph_label", "ph", "type")
  row.names(props) <- NULL
  as_ph_location(props)
}

#' @export
#' @title location of a full size element
#' @description The function will return the location corresponding
#' to a full size display.
#' @param label a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @examples
#' library(magrittr)
#' my_pres <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("Hello world", location = ph_location_fullsize() ) %>%
#'   print(target = tempfile(fileext = ".pptx") )
ph_location_fullsize <- function( label = "", ... ){

  x <- get_doc(...)
  layout_data <- slide_size(x)
  layout_data$left <- 0L
  layout_data$top <- 0L
  layout_data$ph_label <- label
  layout_data$ph <- ""
  layout_data$type <- "body"

  as_ph_location(as.data.frame(layout_data, stringsAsFactors = FALSE))
}

#' @export
#' @title location of a left body element
#' @description The function will return the location corresponding
#' to a left bounding box. The function assume the layout 'Two Content'
#' is existing.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @examples
#' library(magrittr)
#' read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("Hello", location = ph_location_left() ) %>%
#'   ph_with("world", location = ph_location_right() ) %>%
#'   print(target = tempfile(fileext = ".pptx") )
ph_location_left <- function( ... ){

  x <- get_doc(...)
  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm()

  args <- list(...)
  master <- ifelse(is.null(args$master), unique( xfrm$master_name ), args$master)
  get_ph_loc(x, layout = "Two Content", master = master,
             type = "body", position_right = FALSE,
             position_top = TRUE)
}

#' @export
#' @title location of a right body element
#' @description The function will return the location corresponding
#' to a right bounding box. The function assume the layout 'Two Content'
#' is existing.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @examples
#' library(magrittr)
#' read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("Hello", location = ph_location_left() ) %>%
#'   ph_with("world", location = ph_location_right() ) %>%
#'   print(target = tempfile(fileext = ".pptx") )
ph_location_right <- function( ... ){

  x <- get_doc(...)
  slide <- x$slide$get_slide(x$cursor)
  xfrm <- slide$get_xfrm()

  args <- list(...)
  master <- ifelse(is.null(args$master), unique( xfrm$master_name ), args$master)
  get_ph_loc(x, layout = "Two Content", master = master,
             type = "body", position_right = TRUE,
             position_top = TRUE)
}

