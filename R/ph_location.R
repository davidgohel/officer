get_ph_loc <- function(x, layout, master, type, position_right, position_top, id = NULL){

  props <- layout_properties( x, layout = layout, master = master )
  props <- props[props$type %in% type, , drop = FALSE]

  if( nrow(props) < 1) {
    stop("no selected row")
  }
  if( !is.null(id) ){
    props <- props[id,, drop = FALSE]
  } else {
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

#' @export
#' @title eval a location on the current slide
#' @description Eval a shape location against the current slide.
#' This function is to be used to add custom openxml code. A
#' list is returned, it contains informations width, height, left
#' and top positions and other informations necessary to add a
#' content on a slide.
#' @param x a location for a placeholder.
#' @param doc an rpptx object
#' @param ... unused arguments
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content",
#'   master = "Office Theme")
#' fortify_location(ph_location_fullsize(), doc)
#' @seealso \code{\link{ph_location}}, \code{\link{ph_with}}
fortify_location <- function( x, doc, ... ){
  UseMethod("fortify_location")
}

# main ----

#' @export
#' @title create a location for a placeholder
#' @description The function will return a list that complies with
#' expected format for argument \code{location} of function\code{ph_with}.
#' @param left,top,width,height place holder coordinates
#' in inches.
#' @param newlabel a label for the placeholder. See section details.
#' @param bg background color
#' @param rotation rotation angle
#' @param ... unused arguments
#' @family functions for placeholder location
#' @details
#' The location of the bounding box associated to a placeholder
#' within a slide is specified with the left top coordinate,
#' the width and the height. These are defined in inches:
#'
#' \describe{
#'   \item{left}{left coordinate of the bounding box}
#'   \item{top}{top coordinate of the bounding box}
#'   \item{width}{width of the bounding box}
#'   \item{height}{height of the bounding box}
#' }
#'
#' In addition to these attributes, a label can be
#' associated with the shape. Shapes, text boxes, images and other objects
#' will be identified with that label in the Selection Pane of PowerPoint.
#' This label can then be reused by other functions such as `ph_location_label()`.
#' It can be set with argument `newlabel`.
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(doc, "Hello world",
#'   location = ph_location(width = 4, height = 3, newlabel = "hello") )
#' print(doc, target = tempfile(fileext = ".pptx") )
ph_location <- function(left = 1, top = 1, width = 4, height = 3,
                        newlabel = "",
                        bg = NULL, rotation = NULL,
                        ...){

  x <- list(left = left, top = top, width = width, height = height,
    ph_label = newlabel, ph = NA_character_, bg = bg, rotation = rotation)


  class(x) <- c("location_manual", "location_str")
  x
}

#' @export
fortify_location.location_manual <- function( x, doc, ...){
  x
}

#' @title create a location for a placeholder based on a template
#' @description The function will return a list that complies with
#' expected format for argument \code{location} of function
#' \code{ph_with}. A placeholder will be used as template
#' and its positions will be updated with values `left`, `top`, `width`, `height`.
#' @param left,top,width,height place holder coordinates
#' in inches.
#' @param newlabel a label for the placeholder. See section details.
#' @param type placeholder type to look for in the slide layout, one
#' of 'body', 'title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum'.
#' It will be used as a template placeholder.
#' @param id index of the placeholder template. If two body placeholder, there can be
#' two different index: 1 and 2 for the first and second body placeholders defined
#' in the layout.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(doc, "Title",
#'   location = ph_location_type(type = "title") )
#' doc <- ph_with(doc, "Hello world",
#'     location = ph_location_template(top = 4, type = "title") )
#' print(doc, target = tempfile(fileext = ".pptx") )
#' @export
ph_location_template <- function(left = 1, top = 1, width = 4, height = 3,
                        newlabel = "", type = NULL, id = 1,
                        ...){

  x <- list(left = left, top = top, width = width, height = height,
    ph_label = newlabel, ph = NA_character_,
    type = type, id = id)

  class(x) <- c("location_template", "location_str")
  x
}
#' @export
fortify_location.location_template <- function( x, doc, ...){
  slide <- doc$slide$get_slide(doc$cursor)
  if( !is.null( x$type ) ){
    ph <- slide$get_xfrm(type = x$type, index = x$id)$ph
  } else {
    ph <- sprintf('<p:ph type="%s"/>', "body")
  }
  x <- ph_location(left = x$left, top = x$top, width = x$width, height = x$height,
              label = x$ph_label)
  x$ph <- ph
  fortify_location.location_manual(x)
}

#' @export
#' @title location of a placeholder based on a type
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
#' @param id index of the placeholder. If two body placeholder, there can be
#' two different index: 1 and 2 for the first and second body placeholders defined
#' in the layout. If this argument is used, \code{position_right} and \code{position_top}
#' will be ignored.
#' @param newlabel a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#'
#' @example examples/ph_location_type.R
ph_location_type <- function( type = "body", position_right = TRUE, position_top = TRUE, newlabel = NULL, id = NULL, ...){

  ph_types <- c("ctrTitle", "subTitle", "dt", "ftr",
                "sldNum", "title", "body")
  if(!type %in% ph_types){
    stop("argument type must be a value of ", paste0(shQuote(ph_types), collapse = ", ", "."))
  }
  x <- list(type = type, position_right = position_right, position_top = position_top, id = id, label = newlabel)
  class(x) <- c("location_type", "location_str")
  x
}
#' @export
fortify_location.location_type <- function( x, doc, ...){

  slide <- doc$slide$get_slide(doc$cursor)
  xfrm <- slide$get_xfrm()
  args <- list(...)
  layout <- ifelse(is.null(args$layout), unique( xfrm$name ), args$layout)
  master <- ifelse(is.null(args$master), unique( xfrm$master_name ), args$master)
  out <- get_ph_loc(doc, layout = layout, master = master,
             type = x$type, position_right = x$position_right,
             position_top = x$position_top, id = x$id)
  if( !is.null(x$label) )
    out$ph_label <- x$label
  out

}

#' @export
#' @title location of a named placeholder
#' @description The function will use the label of a placeholder
#' to find the corresponding location.
#' @param ph_label placeholder label of the used layout. It can be read in PowerPoint or
#' with function \code{layout_properties()} in column \code{ph_label}.
#' @param newlabel a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#'
#' @example examples/ph_location_label.R
ph_location_label <- function( ph_label, newlabel = NULL, ...){
  x <- list(ph_label = ph_label, label = newlabel)
  class(x) <- c("location_label", "location_str")
  x
}

#' @export
fortify_location.location_label <- function( x, doc, ...){

  slide <- doc$slide$get_slide(doc$cursor)
  xfrm <- slide$get_xfrm()

  layout <- unique( xfrm$name )
  master <- unique(xfrm$master_name)

  props <- layout_properties( doc, layout = layout, master = master )
  props <- props[props$ph_label %in% x$ph_label, , drop = FALSE]

  if( nrow(props) < 1) {
    stop("no selected row")
  }

  if( nrow(props) > 1) {
    warning("more than a row have been selected")
  }

  props <- props[, c("offx", "offy", "cx", "cy", "ph_label", "ph", "type")]
  names(props) <- c("left", "top", "width", "height", "ph_label", "ph", "type")
  row.names(props) <- NULL
  out <- as_ph_location(props)
  if( !is.null(x$label) )
    out$ph_label <- x$label
  out

}

#' @export
#' @title location of a full size element
#' @description The function will return the location corresponding
#' to a full size display.
#' @param newlabel a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(doc, "Hello world", location = ph_location_fullsize() )
#' print(doc, target = tempfile(fileext = ".pptx") )
ph_location_fullsize <- function( newlabel = "", ... ){

  x <- list(label = newlabel)
  class(x) <- c("location_fullsize", "location_str")
  x
}

#' @export
fortify_location.location_fullsize <- function( x, doc, ...){

  layout_data <- slide_size(doc)
  layout_data$left <- 0L
  layout_data$top <- 0L
  if( !is.null(x$label) )
    layout_data$ph_label <- x$label
  layout_data$ph <- NA_character_
  layout_data$type <- "body"

  as_ph_location(as.data.frame(layout_data, stringsAsFactors = FALSE))
}

#' @export
#' @title location of a left body element
#' @description The function will return the location corresponding
#' to a left bounding box. The function assume the layout 'Two Content'
#' is existing.
#' @param newlabel a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(doc, "Hello left", location = ph_location_left() )
#' doc <- ph_with(doc, "Hello right", location = ph_location_right() )
#' print(doc, target = tempfile(fileext = ".pptx") )
ph_location_left <- function( newlabel = NULL, ... ){

  x <- list(label = newlabel)
  class(x) <- c("location_left", "location_str")
  x
}

#' @export
fortify_location.location_left <- function( x, doc, ...){

  slide <- doc$slide$get_slide(doc$cursor)
  xfrm <- slide$get_xfrm()

  args <- list(...)
  master <- if(is.null(args$master)) unique( xfrm$master_name ) else args$master
  out <- get_ph_loc(doc, layout = "Two Content", master = master,
             type = "body", position_right = FALSE,
             position_top = TRUE)
  if( !is.null(x$label) )
    out$ph_label <- x$label
  out
}

#' @export
#' @title location of a right body element
#' @description The function will return the location corresponding
#' to a right bounding box. The function assume the layout 'Two Content'
#' is existing.
#' @param newlabel a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(doc, "Hello left", location = ph_location_left() )
#' doc <- ph_with(doc, "Hello right", location = ph_location_right() )
#' print(doc, target = tempfile(fileext = ".pptx") )
ph_location_right <- function( newlabel = NULL, ... ){

  x <- list(label = newlabel)
  class(x) <- c("location_right", "location_str")
  x
}

#' @export
fortify_location.location_right <- function( x, doc, ...){

  slide <- doc$slide$get_slide(doc$cursor)
  xfrm <- slide$get_xfrm()

  args <- list(...)
  master <- ifelse(is.null(args$master), unique( xfrm$master_name ), args$master)
  out <- get_ph_loc(doc, layout = "Two Content", master = master,
             type = "body", position_right = TRUE,
             position_top = TRUE)
  if( !is.null(x$label) )
    out$ph_label <- x$label
  out
}

