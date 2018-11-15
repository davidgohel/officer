#' @export
#' @title create a location for a placeholder
#' @description The function will return a list that complies with
#' expected format for argument \code{location} of functions \code{ph_with_*}.
#' @param x an rpptx object
#' @param ph_label a label for the placeholder.
#' @family functions for placeholder location
#' @details
#' The location of the bounding box associated to a placeholder
#' within a presentation slide is specified with the left top coordinate,
#' the width and the height. These are defined in inches:
#'
#' \describe{
#'   \item{offx}{left coordinate of the bounding box}
#'   \item{offy}{top coordinate of the bounding box}
#'   \item{width}{width of the bounding box}
#'   \item{height}{height of the bounding box}
#' }
#'
#' In addition to these attributes, there is also an attribute \code{ph_label}
#' that can be associated with the shape (shapes, text boxes, images and other objects
#' will be identified with that label in the Selection Pane of PowerPoint).
#'
#' This result can then be used to position new elements in the slides with
#' functions \code{ph_with_*}.
#' @examples
#' x <- read_pptx()
#' ph_location_fullsize(x)
ph_location <- function(left, top, width, height, label){

  x <- list(
    offx = offx,
    offy = offy,
    width = width,
    height = height,
    ph_label = label
  )
  x
}


#' @export
#' @title location of a placeholder type
#' @description The function will use the type name of the placeholder (e.g. body, title),
#' the layout name and few other criterias to find the corresponding location.
#' @param x an rpptx object
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
#' @param type placeholder type to look for in the slide layout, one
#' of 'body', 'title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum'.
#' @param position_right the parameter is used when a selection with above
#' parameters does not provide a unique position (for example
#' layout 'Two Content' contains two element of type 'body').
#' If \code{TRUE}, the element the most on the right side will be selected,
#' otherwise the element the most on the left side will be selected.
#' @param position_top same than \code{position_right} but applied
#' to top versus bottom.
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' x <- read_pptx()
#' ph_location_type(x)
ph_location_type <- function( x, layout = "Title and Content",
                                master = "Office Theme",
                                type = "body",
                                position_right = TRUE,
                                position_top = TRUE){

  props <- layout_properties( x, layout = layout, master = master )
  props <- props[props$type %in% type, , drop = FALSE]

  if( nrow(props) < 1) {
    stop("no selected row")
  } else if( nrow(props) == 1) {
    return(props)
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
  props <- props[, c("offx", "offy", "cx", "cy", "ph_label", "type")]
  as_ph_location(props)

}

#' @export
#' @title location of a named placeholder
#' @description The function will use the label of a placeholder
#' to find the corresponding location.
#' @param x an rpptx object
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
#' @param ph_label placeholder label. It can be read in PowerPoint or
#' with function \code{layout_properties()} in column \code{ph_label}.
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' x <- read_pptx()
#' ph_location_label(x, layout = "Title and Content",
#'   ph_label = "Content Placeholder 2")
ph_location_label <- function( x, layout = NULL,
                                 master = "Office Theme",
                                 ph_label){

  props <- layout_properties( x, layout = layout, master = master )
  props <- props[props$ph_label %in% ph_label, , drop = FALSE]

  if( nrow(props) < 1) {
    stop("no selected row")
  }

  if( nrow(props) > 1) {
    warning("more than a row have been selected")
  }

  props <- props[, c("offx", "offy", "cx", "cy", "ph_label", "type")]
  row.names(props) <- NULL
  as_ph_location(props)
}

#' @export
#' @title location of a full size element
#' @description The function will return the location corresponding
#' to a full size display.
#' @param x an rpptx object
#' @param label a label to associate with the placeholder.
#' @family functions for placeholder location
#' @examples
#' x <- read_pptx()
#' ph_location_fullsize(x)
ph_location_fullsize <- function( x, label = "" ){
  layout_data <- slide_size(x)
  layout_data$offx <- 0L
  layout_data$offy <- 0L
  layout_data$ph_label <- label
  layout_data$type <- "body"

  as_ph_location(as.data.frame(props, stringsAsFactors = FALSE))
}

#' @export
#' @title location of a left body element
#' @description The function will return the location corresponding
#' to a left bounding box. The function assume the layout 'Two Content'
#' is existing.
#' @param x an rpptx object
#' @family functions for placeholder location
#' @examples
#' x <- read_pptx()
#' ph_location_left(x)
ph_location_left <- function( x ){
  ph_location_type( x, layout = "Two Content",
                                master = "Office Theme",
                                type = "body",
                                position_right = FALSE,
                                position_top = TRUE)
}

#' @export
#' @title location of a right body element
#' @description The function will return the location corresponding
#' to a right bounding box. The function assume the layout 'Two Content'
#' is existing.
#' @param x an rpptx object
#' @family functions for placeholder location
#' @examples
#' x <- read_pptx()
#' ph_location_right(x)
ph_location_right <- function( x ){
  ph_location_type( x, layout = "Two Content",
                                master = "Office Theme",
                                type = "body",
                                position_right = TRUE,
                                position_top = TRUE)
}




as_ph_location <- function(x){
  if( !is.data.frame(x) ){
    stop("x should be a data.frame")
  }
  ref_names <- c( cx = "width", cy = "height", offx = "left", offy = "top",
                 ph_label = "ph_label", "type" = "type")

  if (!(setequal(names(x), names(ref_names)))) {
    stop("missing column values:", paste0(names(ref_names), collapse = ","))
  }

  x <- x[names(ref_names)]
  names(x) <- as.character(ref_names)

  if( x$type %in% c("ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title") ){
    x$ph <- sprintf('<p:ph type="%s"/>', x$type)
  } else x$ph <- ""
  x$type <- NULL
  as.list(x)
}

