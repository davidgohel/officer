
props_to_ph_location <- function(props) {
  if (nrow(props) > 1) {
    cli::cli_alert_warning("More than one placeholder selected.")
  }
  props <- props[, c("offx", "offy", "cx", "cy", "ph_label", "ph", "type", "fld_id", "fld_type", "rotation")]
  names(props) <- c("left", "top", "width", "height", "ph_label", "ph", "type", "fld_id", "fld_type", "rotation")
  as_ph_location(props)
}


# id is deprecated and replaced by type_idx. Will be removed soon
get_ph_loc <- function(x, layout, master, type, type_idx = NULL, position_right, position_top,
                       id = NULL, ph_id = NULL) {
  props <- layout_properties(x, layout = layout, master = master)

  if (!is.null(ph_id)) {
    ids <- sort(stats::na.omit(as.numeric(props$id)))
    if (length(ids) <= 20) {
      .all_ids_switch <- c("x" = "Available ids:  {.val {ids}}.") # only if few ids
    } else {
      .all_ids_switch <- NULL
    }
    if (!ph_id %in% ids) {
      cli::cli_abort(
        c(
          "{.arg id} {.val {ph_id}} does not exist.",
          .all_ids_switch,
          "i" = cli::col_grey("see column {.val id} in {.code layout_properties(..., '{layout}', '{master}')}")
        ),
        call = NULL
      )
    }
    props <- props[props$id == ph_id, , drop = FALSE]
    return(props_to_ph_location(props))
  }

  types_on_layout <- unique(props$type)
  props <- props[props$type %in% type, , drop = FALSE]
  nr <- nrow(props)
  if (nr < 1) {
    cli::cli_abort(c(
      "Found no placeholder of type {.val {type}} on layout {.val {layout}}.",
      "x" = "Available types are {.val {types_on_layout}}",
      "i" = cli::col_grey("see {.code layout_properties(x, '{layout}', '{master}')}")
    ), call = NULL)
  }

  # id and type_idx are both used for now. 'id' is deprecated. The following code block can be removed in the future.
  if (!is.null(id)) {
    if (!id %in% 1L:nr) {
      cli::cli_abort(
        c(
          "{.arg id} is out of range.",
          "x" = "Must be between {.val {1L}} and {.val {nr}} for ph type {.val {type}}.",
          "i" = cli::col_grey("see {.code layout_properties(x, '{layout}', '{master}')} for all phs with type '{type}'")
        ),
        call = NULL
      )
    }
    # the ordering of 'type_idx' (top->bottom, left-righ) is different than for the 'id' arg (index
    # along the id colomn). Here, we restore the old ordering, to avoid a breaking change.
    props <- props[order(props$type, as.integer(props$id)), ] # set order for type idx. Removing the line would result in the default layout properties order, i.e., top->bottom left->right.
    props$.id <- stats::ave(props$type, props$master_name, props$name, props$type, FUN = seq_along)
    props <- props[props$.id == id, , drop = FALSE]
    return(props_to_ph_location(props))
  }

  if (!is.null(type_idx)) {
    if (!type_idx %in% props$type_idx) {
      cli::cli_abort(
        c(
          "{.arg type_idx} is out of range.",
          "x" = "Must be between {.val {1L}} and {.val {max(props$type_idx)}} for ph type {.val {type}}.",
          "i" = cli::col_grey("see {.code layout_properties(..., layout = '{layout}', master = '{master}')} for indexes of type '{type}'")
        ),
        call = NULL
      )
    }
    props <- props[props$type_idx == type_idx, , drop = FALSE]
    return(props_to_ph_location(props))
  }

  if (position_right) {
    props <- props[props$offx + 0.0001 > max(props$offx), ]
  } else {
    props <- props[props$offx - 0.0001 < min(props$offx), ]
  }
  if (position_top) {
    props <- props[props$offy - 0.0001 < min(props$offy), ]
  } else {
    props <- props[props$offy + 0.0001 > max(props$offy), ]
  }
  props_to_ph_location(props)
}


as_ph_location <- function(x, ...) {
  if (!is.data.frame(x)) {
    cli::cli_abort(
      c("{.arg x} must be a data frame.",
        "x" =  "You provided {.cls {class(x)[1]}} instead.")
      )
  }
  ref_names <- c(
    "width", "height", "left", "top", "ph_label", "ph", "type", "rotation", "fld_id", "fld_type"
  )
  if (!all(is.element(ref_names, names(x)))) {
    stop("missing column values:", paste0(setdiff(ref_names, names(x)), collapse = ","))
  }
  out <- x[ref_names]
  as.list(out)
}


#' @export
#' @title Eval a location on the current slide
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
#' @family functions for officer extensions
#' @keywords internal
fortify_location <- function( x, doc, ... ){
  UseMethod("fortify_location")
}

# main ----

#' @export
#' @title Location for a placeholder from scratch
#' @description The function will return a list that complies with
#' expected format for argument \code{location} of function \code{ph_with}.
#' @param left,top,width,height place holder coordinates
#' in inches.
#' @param newlabel a label for the placeholder. See section details.
#' @param bg background color
#' @param rotation rotation angle
#' @param ln a [sp_line()] object specifying the outline style.
#' @param geom shape geometry, see http://www.datypic.com/sc/ooxml/t-a_ST_ShapeType.html
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
#'
#' # Set geometry and outline
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' loc <- ph_location(left = 1, top = 1, width = 4, height = 3, bg = "steelblue",
#'                    ln = sp_line(color = "red", lwd = 2.5),
#'                    geom = "trapezoid")
#' doc <- ph_with(doc, "", loc = loc)
#' print(doc, target = tempfile(fileext = ".pptx") )
ph_location <- function(left = 1, top = 1, width = 4, height = 3,
                        newlabel = "",
                        bg = NULL, rotation = NULL,
                        ln = NULL, geom = NULL,
                        ...){

  x <- list(left = left, top = top, width = width, height = height,
    ph_label = newlabel, ph = NA_character_, bg = bg, rotation = rotation, ln = ln, geom = geom, fld_type = NA_character_, fld_id = NA_character_)

  class(x) <- c("location_manual", "location_str")
  x
}

#' @export
fortify_location.location_manual <- function( x, doc, ...){
  x
}

#' @title Location for a placeholder based on a template
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
#' @title Location of a placeholder based on a type
#' @description The function will use the type name of the placeholder (e.g. body, title),
#' the layout name and few other criterias to find the corresponding location.
#' @param type placeholder type to look for in the slide layout, one
#' of 'body', 'title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum'.
#' @param position_right the parameter is used when a selection with above
#' parameters does not provide a unique position (for example
#' layout 'Two Content' contains two element of type 'body').
#' If `TRUE`, the element the most on the right side will be selected,
#' otherwise the element the most on the left side will be selected.
#' @param position_top same than `position_right` but applied
#' to top versus bottom.
#' @param type_idx Type index of the placeholder. If there is more than one
#' placeholder of a type (e.g., `body`), the type index can be supplied to uniquely
#' identify a ph. The index is a running number starting at 1. It is assigned by
#' placeholder position  (top -> bottom, left -> right). See [plot_layout_properties()]
#' for details. If `idx` argument is used, `position_right` and `position_top`
#' are ignored.
#' @param id (**DEPRECATED, use `type_idx` instead**) Index of the placeholder.
#' If two body placeholder, there can be two different index: 1 and 2 for the
#' first and second body placeholders defined in the layout. If this argument
#' is used, `position_right` and `position_top` will be ignored.
#' @param newlabel a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' # ph_location_type demo ----
#'
#' loc_title <- ph_location_type(type = "title")
#' loc_footer <- ph_location_type(type = "ftr")
#' loc_dt <- ph_location_type(type = "dt")
#' loc_slidenum <- ph_location_type(type = "sldNum")
#' loc_body <- ph_location_type(type = "body")
#'
#'
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, "Un titre", location = loc_title)
#' doc <- ph_with(x = doc, "pied de page", location = loc_footer)
#' doc <- ph_with(x = doc, format(Sys.Date()), location = loc_dt)
#' doc <- ph_with(x = doc, "slide 1", location = loc_slidenum)
#' doc <- ph_with(x = doc, letters[1:10], location = loc_body)
#'
#' loc_subtitle <- ph_location_type(type = "subTitle")
#' loc_ctrtitle <- ph_location_type(type = "ctrTitle")
#' doc <- add_slide(doc, layout = "Title Slide", master = "Office Theme")
#' doc <- ph_with(x = doc, "Un sous titre", location = loc_subtitle)
#' doc <- ph_with(x = doc, "Un titre", location = loc_ctrtitle)
#'
#' fileout <- tempfile(fileext = ".pptx")
#' print(doc, target = fileout)
#'
ph_location_type <- function(type = "body", type_idx = NULL, position_right = TRUE, position_top = TRUE,
                             newlabel = NULL, id = NULL, ...) {
  # the following two warnings can be deleted after the deprecated id arg is removed.
  if (!is.null(id) && !is.null(type_idx)) {
    cli::cli_warn("{.arg id} is ignored if {.arg type_idx} is provided ")
  }
  if (!is.null(id) && is.null(type_idx)) {
    cli::cli_warn(
      c(
        "!" = "The {.arg id} argument in {.fn ph_location_type} is deprecated as of {.pkg officer} 0.6.7.",
        "i" = "Please use the {.arg type_idx} argument instead.",
        "x" = cli::col_red("Caution: new index logic in {.arg type_idx} (see docs).")
      )
    )
  }

  ph_types <- c(
    "ctrTitle", "subTitle", "dt", "ftr", "sldNum", "title", "body",
    "pic", "chart", "tbl", "dgm", "media", "clipArt"
  )
  if (!type %in% ph_types) {
    cli::cli_abort(
      c("type {.val {type}} is unknown.",
        "x" = "Must be one of {.or {.val {ph_types}}}"
      ),
      call = NULL
    )
  }
  x <- list(
    type = type, type_idx = type_idx, position_right = position_right,
    position_top = position_top, id = id, label = newlabel
  )
  class(x) <- c("location_type", "location_str")
  x
}


#' @export
fortify_location.location_type <- function(x, doc, ...) {
  slide <- doc$slide$get_slide(doc$cursor)
  xfrm <- slide$get_xfrm()
  args <- list(...)

  layout <- ifelse(is.null(args$layout), unique(xfrm$name), args$layout)
  master <- ifelse(is.null(args$master), unique(xfrm$master_name), args$master)

  # to avoid a breaking change, the deprecated id is passed along.
  # As type_idx uses a different index order than id, this is necessary until the id arg is removed.
  out <- get_ph_loc(doc,
    layout = layout, master = master,
    type = x$type, position_right = x$position_right,
    position_top = x$position_top, type_idx = x$type_idx,
    id = x$id, ph_id = NULL # id is deprecated and will be removed soon
  )
  if (!is.null(x$label)) {
    out$ph_label <- x$label
  }
  out
}


#' @export
#' @title Location of a named placeholder
#' @description The function will use the label of a placeholder
#' to find the corresponding location.
#' @param ph_label placeholder label of the used layout. It can be read in PowerPoint or
#' with function \code{layout_properties()} in column \code{ph_label}.
#' @param newlabel a label to associate with the placeholder.
#' @param ... unused arguments
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' # ph_location_label demo ----
#'
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content")
#'
#' # all ph_label can be read here
#' layout_properties(doc, layout = "Title and Content")
#'
#' doc <- ph_with(doc, head(iris),
#'   location = ph_location_label(ph_label = "Content Placeholder 2") )
#' doc <- ph_with(doc, format(Sys.Date()),
#'   location = ph_location_label(ph_label = "Date Placeholder 3") )
#' doc <- ph_with(doc, "This is a title",
#'   location = ph_location_label(ph_label = "Title 1") )
#'
#' print(doc, target = tempfile(fileext = ".pptx"))
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
    stop("Placeholder ", shQuote(x$ph_label),
         " in the slide layout is duplicated. It needs to be unique. Hint: layout_dedupe_ph_labels() helps handling duplicates.")
  }

  props <- props[, c("offx", "offy", "cx", "cy", "ph_label", "ph", "type", "rotation", "fld_id", "fld_type")]
  names(props) <- c("left", "top", "width", "height", "ph_label", "ph", "type", "rotation", "fld_id", "fld_type")
  row.names(props) <- NULL
  out <- as_ph_location(props)
  if( !is.null(x$label) )
    out$ph_label <- x$label
  out

}


#' @export
#' @title Location of a full size element
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
  layout_data$rotation <- 0L
  layout_data$fld_id <- NA_character_
  layout_data$fld_type <- NA_character_

  as_ph_location(as.data.frame(layout_data, stringsAsFactors = FALSE))
}

#' @export
#' @title Location of a left body element
#' @description The function will return the location corresponding
#' to a left bounding box. The function assume the layout 'Two Content'
#' is existing. This is an helper function, if you don't have a layout
#' named 'Two Content', use [ph_location_type()] and set arguments
#' to your specific needs.
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
#' @title Location of a right body element
#' @description The function will return the location corresponding
#' to a right bounding box. The function assume the layout 'Two Content'
#' is existing. This is an helper function, if you don't have a layout
#' named 'Two Content', use [ph_location_type()] and set arguments
#' to your specific needs.
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


#' @export
#' @title Location of a placeholder based on its id
#' @description Each placeholder has an id (a low integer value). The ids are unique across a single
#' layout. The function uses the placeholder's id to reference it. Different from a ph label,
#' the id is auto-assigned by PowerPoint and cannot be modified by the user.
#' Use [layout_properties()] (column `id`) and [plot_layout_properties()] (upper right
#' corner, in green) to find a placeholder's id.
#'
#' @param id placeholder id.
#' @param newlabel a new label to associate with the placeholder.
#' @param ... not used.
#' @family functions for placeholder location
#' @inherit ph_location details
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc, "Comparison")
#' plot_layout_properties(doc, "Comparison")
#'
#' doc <- ph_with(doc, "The Title", location = ph_location_id(id = 2)) # title
#' doc <- ph_with(doc, "Left Header", location = ph_location_id(id = 3)) # left header
#' doc <- ph_with(doc, "Left Content", location = ph_location_id(id = 4)) # left content
#' doc <- ph_with(doc, "The Footer", location = ph_location_id(id = 8)) # footer
#'
#' file <- tempfile(fileext = ".pptx")
#' print(doc, file)
#' \dontrun{
#' file.show(file) # may not work on your system
#' }
ph_location_id <- function(id, newlabel = NULL, ...) {
  ph_id <- id # for disambiguation, store initial value

  if (length(ph_id) > 1) {
    cli::cli_abort(
      c("{.arg id} must be {cli::style_underline('one')} number",
        "x" = "Found more than one entry: {.val {ph_id}}"
      )
    )
  }
  if (is.null(ph_id) || is.na(ph_id) || length(ph_id) == 0) {
    cli::cli_abort("{.arg id} must be a positive number")
  }
  if (!is.integer(ph_id)) {
    ph_id <- suppressWarnings(as.integer(ph_id))
    if (is.na(ph_id)) {
      cli::cli_abort(
        c("Cannot convert {.val {id}} to integer",
          "x" = "{.arg id} must be a number, you provided class {.cls {class(id)[1]}}"
        )
      )
    }
  }
  if (ph_id < 1) {
    cli::cli_abort(
      c("{.arg id} must be a {cli::style_underline('positive')} number",
        "x" = "Found {.val {ph_id}}"
      )
    )
  }
  x <- list(
    type = NULL, type_idx = NULL, position_right = NULL, position_right = NULL,
    position_top = NULL, id = NULL, ph_id = ph_id, label = newlabel
  )
  class(x) <- c("location_id", "location_num")
  x
}


#' @export
fortify_location.location_id <- function(x, doc, ...) {
  slide <- doc$slide$get_slide(doc$cursor)
  xfrm <- slide$get_xfrm()
  args <- list(...)

  layout <- ifelse(is.null(args$layout), unique(xfrm$name), args$layout)
  master <- ifelse(is.null(args$master), unique(xfrm$master_name), args$master)
  out <- get_ph_loc(doc, layout = layout, master = master, ph_id = x$ph_id)
  if (!is.null(x$label)) {
    out$ph_label <- x$label
  }
  out
}


