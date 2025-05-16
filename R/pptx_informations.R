#' @export
#' @title Number of slides
#' @description Function `length` will return the number of slides.
#' @param x an rpptx object
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres, "Title and Content")
#' my_pres <- add_slide(my_pres, "Title and Content")
#' length(my_pres)
#' @family functions for reading presentation information
length.rpptx <- function( x ){
  x$slide$length()
}


#' @export
#' @title Slides width and height
#' @description Get the width and height of slides in inches as
#' a named vector.
#' @inheritParams length.rpptx
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres,
#'   layout = "Two Content", master = "Office Theme")
#' slide_size(my_pres)
#' @family functions for reading presentation information
slide_size <- function(x) {
  pres <- x$presentation$get()
  dimensions <- xml_attrs(xml_find_first(pres, "p:sldSz"))
  dimensions <- as.list(as.integer(dimensions[c("cx", "cy")]) / 914400)
  names(dimensions) <- c("width", "height")
  dimensions
}


#' @export
#' @title Presentation layouts summary
#' @description Get information about slide layouts and
#' master layouts into a data.frame. This function returns
#' a data.frame containing all layout and master names.
#' @inheritParams length.rpptx
#' @examples
#' my_pres <- read_pptx()
#' layout_summary ( x = my_pres )
#' @family functions for reading presentation information
layout_summary <- function( x ){
  data <- x$slideLayouts$get_metadata()
  data.frame(layout = data$name, master = data$master_name, stringsAsFactors = FALSE)
}


#' @export
#' @title Slide layout properties
#' @description Detailed information about the placeholders on the slide layouts (label, position, etc.).
#' See *Value* section below for more info.
#' @param x an `rpptx` object
#' @param layout slide layout name. If `NULL`, returns all layouts.
#' @param master master layout name where `layout` is located. If `NULL`, returns all masters.
#' @returns Returns a data frame with one row per placeholder and the following columns:
#' * `master_name`: Name of master (a `.pptx` file may have more than one)
#' * `name`: Name of layout
#' * `type`: Placeholder type
#' * `type_idx`: Running index for phs of the same type. Ordering by ph position
#'    (top -> bottom, left -> right)
#' * `id`: A unique placeholder id (assigned by PowerPoint automatically, starts at 2, potentially non-consecutive)
#' * `ph_label`: Placeholder label (can be set by the user in PowerPoint)
#' * `ph`: Placholder XML fragment (usually not needed)
#' * `offx`,`offy`: placeholder's distance from left and top edge (in inch)
#' * `cx`,`cy`: width and height of placeholder (in inch)
#' * `rotation`: rotation in degrees
#' * `fld_id` is generally stored as a hexadecimal or GUID value
#' * `fld_type`: a unique identifier for a particular field
#'
#' @examples
#' x <- read_pptx()
#' layout_properties(x = x, layout = "Title Slide", master = "Office Theme")
#' layout_properties(x = x, master = "Office Theme")
#' layout_properties(x = x, layout = "Two Content")
#' layout_properties(x = x)
#' @family functions for reading presentation information
layout_properties <- function(x, layout = NULL, master = NULL) {
  data <- x$slideLayouts$get_xfrm_data()
  if (!is.null(layout) && !is.null(master)) {
    data <- data[data$name == layout & data$master_name %in% master, ]
  } else if (is.null(layout) && !is.null(master)) {
    data <- data[data$master_name %in% master, ]
  } else if (!is.null(layout) && is.null(master)) {
    data <- data[data$name == layout, ]
  }
  data <- data[, c("master_name", "name", "type", "type_idx", "id", "ph_label", "ph",
                   "offx", "offy", "cx", "cy", "rotation", "fld_id", "fld_type")]
  data[["offx"]] <- data[["offx"]] / 914400
  data[["offy"]] <- data[["offy"]] / 914400
  data[["cx"]] <- data[["cx"]] / 914400
  data[["cy"]] <- data[["cy"]] / 914400
  data[["rotation"]] <- data[["rotation"]] / 60000

  rownames(data) <- NULL
  data
}


#' @export
#' @title Slide layout properties plot
#' @description Plot slide layout properties into corresponding placeholders.
#'  This can be useful to help visualize placeholders locations and identifiers.
#'  *All* information in the plot stems from the [layout_properties()] output.
#'  See *Details* section for more info.
#' @details
#' The plot contains all relevant information to reference a placeholder via the `ph_location_*`
#' function family:
#'
#' * `label`: ph label (red, center) to be used in [ph_location_label()].
#'    _NB_: The label can be assigned by the user in PowerPoint.
#' * `type[idx]`: ph type + type index in brackets (blue, upper left) to be used in [ph_location_type()].
#'    _NB_: The index is consecutive and is sorted by ph position (top -> bottom, left -> right).
#' * `id`: ph id (green, upper right) to be used in `ph_location_id()` (forthcoming).
#'    _NB_: The id is set by PowerPoint automatically and lack a meaningful order.
#'
#' @param x an `rpptx` object
#' @param layout slide layout name or numeric index (row index from [layout_summary()]. If `NULL` (default), it plots
#'   the current slide's layout or the default layout (if set and there are not slides yet).
#' @param master master layout name where `layout` is located. Can be omitted if layout is unambiguous.
#' @param title if `TRUE` (default), adds a title with the layout name at the top.
#' @param labels if `TRUE` (default), adds placeholder labels (centered in *red*).
#' @param type if `TRUE` (default), adds the placeholder type and its index (in square brackets)
#'   in the upper left corner (in *blue*).
#' @param id if `TRUE` (default), adds the placeholder's unique `id` (see column `id` from
#'  [layout_properties()]) in the upper right corner (in *green*).
#' @param cex List or vector to specify font size for `labels`, `type`, and `id`. Default is
#'   `c(labels = .5, type = .5, id = .5)`. See [graphics::text()] for details on how `cex` works.
#'   Matching by position and partial name matching is supported. A single numeric value will apply to all
#'   three parameters.
#' @param legend Add a legend to the plot (default `FALSE`).
#' @importFrom graphics plot rect text box
#' @family functions for reading presentation information
#' @example inst/examples/example_plot_layout_properties.R
#'
plot_layout_properties <- function(x, layout = NULL, master = NULL, labels = TRUE, title = TRUE, type = TRUE,
                                   id = TRUE, cex = c(labels = .5, type = .5, id = .5), legend = FALSE) {
  stop_if_not_rpptx(x, "x")
  loffset <- ifelse(legend, 1, 0) # make space for legend at top
  old_par <- par(mar = c(2, 2, 1.5 + loffset, 0))
  on.exit(par(old_par))

  .cex <- update_named_defaults(cex, default = list(labels = .5, type = .5, id = .5), default_if_null = TRUE, argname = "cex")
  if (.cex$labels <= 0) labels <- FALSE
  if (.cex$type <= 0) type <- FALSE
  if (.cex$id <= 0) id <- FALSE

  if (is.null(layout) && is.null(master) && length(x) == 0) { # fail or use default layout if set
    if (!has_layout_default(x)) {
      cli::cli_abort(
        c("No {.arg layout} selected and no slides in presentation.",
          "x" = "Pass a layout name or index (see {.fn layout_summary})"
        )
      )
    }
    .ld <- get_layout_default(x)
    la <- get_layout(x, layout = .ld$layout, master = .ld$master)
    cli::cli_inform(c("i" = "Showing default layout: {.val {la$layout_name}}"))
  } else if (is.null(layout) && is.null(master) && length(x) > 0) { # use current slides layout as default (if layout and master are NULL)
    la <- get_layout_for_current_slide(x)
    cli::cli_inform(c("i" = "Showing current slide's layout: {.val {la$layout_name}}"))
  } else {
    la <- get_layout(x, layout, master, layout_by_id = TRUE)
  }

  dat <- layout_properties(x, layout = la$layout_name, master = la$master_name)

  s <- slide_size(x)
  h <- s$height
  w <- s$width
  offx <- offy <- cx <- cy <- NULL # avoid R CMD CHECK problem
  list2env(dat[, c("offx", "offy", "cx", "cy")], environment()) # make available inside functions

  plot(x = c(0, w), y = -c(0, h), asp = 1, type = "n", axes = FALSE, xlab = NA, ylab = NA)
  rect(xleft = 0, xright = w, ybottom = 0, ytop = -h, border = "darkgrey")
  rect(xleft = offx, xright = offx + cx, ybottom = -offy, ytop = -(offy + cy))
  mtext("y [inch]", side = 2, line = 0, cex = .9, col = "darkgrey")
  mtext("x [inch]", side = 1, line = 0, cex = .9, col = "darkgrey")

  if (title) {
    title(main = paste("Layout:", la$layout_name), line = 0 + loffset)
  }
  if (labels) { # centered
    text(x = offx + cx / 2, y = -(offy + cy / 2), labels = dat$ph_label, cex = .cex$labels, col = "red", adj = c(.5, 1)) # adj-vert: avoid interference with type/id in small phs
  }
  if (type) { # upper left corner
    .type_info <- paste0(dat$type, " [", dat$type_idx, "]") # type + index in brackets
    text(x = offx, y = -offy, labels = .type_info, cex = .cex$type, col = "blue", adj = c(-.1, 1.2))
  }
  if (id) { # upper right corner
    text(x = offx + cx, y = -offy, labels = dat$id, cex = .cex$id, col = "darkgreen", adj = c(1.3, 1.2))
  }
  if (legend) {
    legend(
      x = w / 2, y = 0, x.intersp = 0.4, xjust = .5, yjust = 0,
      legend = c("type [type_idx]", "ph_label", "id"), fill = c("blue", "red", "darkgreen"),
      bty = "n", pt.cex = 1.2, cex = .7, text.width = NA,
      text.col = c("blue", "red", "darkgreen"), horiz = TRUE, xpd = TRUE
    )
  }
}


#' @export
#' @title Placeholder parameters annotation
#' @description generates a slide from each layout in the base document to
#' identify the placeholder indexes, types, names, master names and layout names.
#'
#' This is to be used when need to know what parameters should be used with
#' `ph_location*` calls. The parameters are printed in their corresponding shapes.
#'
#' Note that if there are duplicated `ph_label`, you should not use `ph_location_label()`.
#' Hint: You can dedupe labels using [layout_dedupe_ph_labels()].
#'
#' @param path path to the pptx file to use as base document or NULL to use the officer default
#' @param output_file filename to store the annotated powerpoint file or NULL to suppress generation
#' @return rpptx object of the annotated PowerPoint file
#' @examples
#' # To generate an anotation of the default base document with officer:
#' annotate_base(output_file = tempfile(fileext = ".pptx"))
#'
#' # To generate an annotation of the base document 'mydoc.pptx' and place the
#' # annotated output in 'mydoc_annotate.pptx'
#' # annotate_base(path = 'mydoc.pptx', output_file='mydoc_annotate.pptx')
#'
#' @family functions for reading presentation information
annotate_base <- function(path = NULL, output_file = 'annotated_layout.pptx' ){
  ppt <- read_pptx(path=path)
  while(length(ppt)>0){
    ppt <- remove_slide(ppt, 1)
  }

  # Pulling out all of the layouts stored in the template
  lay_sum <- layout_summary(ppt)

  # Looping through each layout
  for(lidx in seq_len(nrow(lay_sum))){
    # Pulling out the layout properties
    layout <- lay_sum[lidx, 1]
    master <- lay_sum[lidx, 2]
    lp <- layout_properties ( x = ppt, layout = layout, master = master)

    # Adding a slide for the current layout
    ppt <- add_slide(x=ppt, layout = layout, master = master)
    size <- slide_size(ppt)
    fpar_ <- fpar(sprintf('layout ="%s", master = "%s"', layout, master),
                  fp_t = fp_text(color = "orange", font.size = 20),
                  fp_p = fp_par(text.align = "right", padding = 5)
    )
    ppt <- ph_with(x = ppt, value = fpar_, ph_label = "layout_ph",
                   location = ph_location(left = 0, top = -0.5, width = size$width, height = 1,
                                          bg = "transparent", newlabel = "layout_ph"))

    # Blank slides have nothing
    if(length(lp[,1] > 0)){
      # Now we go through each placholder
      for(pidx in seq_len(nrow(lp))){
        textstr <- paste("type=", lp$type[pidx], ", index=", lp$id[pidx], ", ph_label=",lp$ph_label[pidx])
        ppt <- ph_with(x=ppt,  value = textstr, location = ph_location_label(type = lp$type[pidx], ph_label = lp$ph_label[pidx]))
      }
    }
  }

  if(!is.null(output_file)){
    print(ppt, target = output_file)
  }

  ppt
}


#' @export
#' @title Slide content in a data.frame
#' @description Get content and positions of current slide
#' into a data.frame. Data for any tables, images, or paragraphs are
#' imported into the resulting data.frame.
#' @note
#' The column `id` of the result is not to be used by users.
#' This is a technical string id whose value will be used by office
#' when the document will be rendered. This is not related to argument
#' `index` required by functions [ph_with()].
#' @inheritParams length.rpptx
#' @param index slide index
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres, "Title and Content")
#' my_pres <- ph_with(my_pres, format(Sys.Date()),
#'   location = ph_location_type(type="dt"))
#' my_pres <- add_slide(my_pres, "Title and Content")
#' my_pres <- ph_with(my_pres, iris[1:2,],
#'   location = ph_location_type(type="body"))
#' slide_summary(my_pres)
#' slide_summary(my_pres, index = 1)
#' @family functions for reading presentation information
slide_summary <- function( x, index = NULL ){

  l_ <- length(x)
  if( l_ < 1 ){
    stop("presentation contains no slide", call. = FALSE)
  }

  if( is.null(index) )
    index <- x$cursor

  if( !between(index, 1, l_ ) ){
    stop("unvalid index ", index, " (", l_," slide(s))", call. = FALSE)
  }

  slide <- x$slide$get_slide(index)

  nodes <- xml_find_all(slide$get(), as_xpath_content_sel("p:cSld/p:spTree/") )
  data <- read_xfrm(nodes, file = "slide", name = "" )
  data$text <- sapply(nodes, xml_text )
  data[["offx"]] <- data[["offx"]] / 914400
  data[["offy"]] <- data[["offy"]] / 914400
  data[["cx"]] <- data[["cx"]] / 914400
  data[["cy"]] <- data[["cy"]] / 914400

  data$name <- NULL
  data$file <- NULL
  data$ph <- NULL
  data
}


#' @export
#' @title Color scheme of a PowerPoint file
#' @description Get the color scheme of a
#' 'PowerPoint' master layout into a data.frame.
#' @inheritParams length.rpptx
#' @examples
#' x <- read_pptx()
#' color_scheme ( x = x )
#' @family functions for reading presentation information
color_scheme <- function( x ){
  x$masterLayouts$get_color_scheme()
}
