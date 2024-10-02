#' @export
#' @title Add a slide
#' @description Add a slide into a pptx presentation.
#' @param x an rpptx object
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
#' @examples
#' my_pres <- read_pptx()
#' layout_summary(my_pres)
#' my_pres <- add_slide(my_pres,
#'   layout = "Two Content", master = "Office Theme"
#' )
#' @seealso [print.rpptx()], [read_pptx()], [plot_layout_properties()], [ph_with()], [layout_summary()]
#' @family functions slide manipulation
add_slide <- function(x, layout = "Title and Content", master = "Office Theme") {
  slide_info <- x$slideLayouts$get_metadata()
  slide_info <- slide_info[slide_info$name == layout & slide_info$master_name == master, ]

  if (nrow(slide_info) < 1) {
    stop("could not find layout named ", shQuote(layout), " in master named ", shQuote(master))
  } else if (nrow(slide_info) > 1) {
    stop(
      "found two layouts named ", shQuote(layout), " in master named ", shQuote(master),
      ". Layout names should not be duplicated."
    )
  }
  new_slidename <- x$slide$get_new_slidename()

  xml_file <- file.path(x$package_dir, "ppt/slides", new_slidename)
  layout_obj <- x$slideLayouts$collection_get(slide_info$filename)
  layout_obj$write_template(xml_file)

  # update presentation elements
  x$presentation$add_slide(target = file.path("slides", new_slidename))
  x$content_type$add_slide(partname = file.path("/ppt/slides", new_slidename))

  x$slide$add_slide(xml_file, x$slideLayouts$get_xfrm_data())

  x$cursor <- x$slide$length()
  x
}


#' @export
#' @title Change current slide
#' @description Change current slide index of an rpptx object.
#' @param x an rpptx object
#' @param index slide index
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- on_slide(doc, index = 1)
#' doc <- ph_with(
#'   x = doc, "First title",
#'   location = ph_location_type(type = "title")
#' )
#' doc <- on_slide(doc, index = 3)
#' doc <- ph_with(
#'   x = doc, "Third title",
#'   location = ph_location_type(type = "title")
#' )
#'
#' file <- tempfile(fileext = ".pptx")
#' print(doc, target = file)
#' @family functions slide manipulation
#' @seealso [read_pptx()], [ph_with()]
on_slide <- function(x, index) {
  l_ <- length(x)
  if (l_ < 1) {
    stop("presentation contains no slide", call. = FALSE)
  }
  if (!between(index, 1, l_)) {
    stop("unvalid index ", index, " (", l_, " slide(s))", call. = FALSE)
  }

  filename <- basename(x$presentation$slide_data()$target[index])
  location <- which(x$slide$get_metadata()$name %in% filename)

  x$cursor <- x$slide$slide_index(filename)

  x
}

#' @export
#' @title Remove a slide
#' @description Remove a slide from a pptx presentation.
#' @param x an rpptx object
#' @param index slide index, default to current slide position.
#' @param rm_images if TRUE (defaults to FALSE), images presented in
#' the slide to remove are also removed from the file.
#' @note cursor is set on the last slide.
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres)
#' my_pres <- remove_slide(my_pres)
#' @family functions slide manipulation
#' @seealso [read_pptx()], [ph_with()], [ph_remove()]
remove_slide <- function(x, index = NULL, rm_images = FALSE) {
  l_ <- length(x)
  if (l_ < 1) {
    stop("presentation contains no slide to delete", call. = FALSE)
  }

  if (is.null(index)) {
    index <- x$cursor
  }

  if (!between(index, 1, l_)) {
    stop("unvalid index ", index, " (", l_, " slide(s))", call. = FALSE)
  }
  filename <- basename(x$presentation$slide_data()$target[index])
  location <- which(x$slide$get_metadata()$name %in% filename)

  if (rm_images) {
    media_files <- pptx_summary(x)$media_file
    if (length(media_files)) {
      unlink(file.path(x$package_dir, media_files), force = TRUE)
    }
  }

  del_file <- x$slide$remove_slide(location)

  # update presentation elements
  x$presentation$remove_slide(del_file)
  x$content_type$remove_slide(partname = del_file)
  x$cursor <- x$slide$length()
  x
}



#' @export
#' @title Move a slide
#' @description Move a slide in a pptx presentation.
#' @inheritParams remove_slide
#' @param to new slide index.
#' @note cursor is set on the last slide.
#' @examples
#' x <- read_pptx()
#' x <- add_slide(x)
#' x <- ph_with(x, "Hello world 1", location = ph_location_type())
#' x <- add_slide(x)
#' x <- ph_with(x, "Hello world 2", location = ph_location_type())
#' x <- move_slide(x, index = 1, to = 2)
#' @family functions slide manipulation
#' @seealso [read_pptx()]
move_slide <- function(x, index = NULL, to) {
  x$presentation$slide_data()

  if (is.null(index)) {
    index <- x$cursor
  }

  l_ <- length(x)

  if (l_ < 1) {
    stop("presentation contains no slide", call. = FALSE)
  }
  if (!between(index, 1, l_)) {
    stop("unvalid index ", index, " (", l_, " slide(s))", call. = FALSE)
  }
  if (!between(to, 1, l_)) {
    stop("unvalid 'to' ", to, " (", l_, " slide(s))", call. = FALSE)
  }

  x$presentation$move_slide(from = index, to = to)
  x$cursor <- to
  x
}



#' @title Correct pptx content references
#' @description Content references are not managed directly
#' but computed after the content is added. This function
#' loop over each slide and fix references (links and images)
#' if necessary.
#' @param x an rpptx object
#' @return an rpptx object
#' @noRd
pptx_fortify_slides <- function(x) {
  for (cursor_index in seq_len(x$slide$length())) {
    slide <- x$slide$get_slide(cursor_index)

    process_images(slide, slide$relationship(), x$package_dir, media_dir = "ppt/media", media_rel_dir = "../media")

    process_links(slide, type = "pml")

    cnvpr <- xml_find_all(slide$get(), "//p:cNvPr")
    for (i in seq_along(cnvpr)) {
      xml_attr(cnvpr[[i]], "id") <- i
    }
  }

  x
}

#' @export
#' @title pptx tags for visual and non visual properties
#' @description Visual and non visual properties of a
#' shape can be returned by this function.
#' @param left,top,width,height place holder coordinates
#' in inches.
#' @param bg background color
#' @param rot rotation angle
#' @param label a label for the placeholder.
#' @param ph string containing xml code for ph tag
#' @param ln a [sp_line()] object specifying the outline style.
#' @param geom shape geometry, see http://www.datypic.com/sc/ooxml/t-a_ST_ShapeType.html
#' @return a character value
#' @family functions for officer extensions
#' @keywords internal
shape_properties_tags <- function(left = 0, top = 0, width = 3, height = 3,
                                  bg = "transparent", rot = 0, label = "", ph = "<p:ph/>",
                                  ln = sp_line(lwd = 0, linecmpd = "solid", lineend = "rnd"), geom = NULL) {
  if (!is.null(bg) && !is.color(bg)) {
    stop("bg must be a valid color.", call. = FALSE)
  }

  bg_str <- solid_fill_pml(bg)
  ln_str <- ln_pml(ln)
  geom_str <- prst_geom_pml(geom)

  xfrm_str <- a_xfrm_str(left = left, top = top, width = width, height = height, rot = rot)
  if (is.null(ph) || is.na(ph)) {
    ph <- "<p:ph/>"
  }

  randomid <- officer::uuid_generate()

  str <- "<p:nvSpPr><p:cNvPr id=\"%s\" name=\"%s\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr>%s%s%s%s</p:spPr>"

  sprintf(str, randomid, label, ph, xfrm_str, geom_str, bg_str, ln_str)
}


# check if slide index exists
ensure_slide_index_exists <- function(x, slide_idx) {
  stop_if_not_rpptx(x)
  if (!is.numeric(slide_idx)) {
    cli::cli_abort(
      c("{.arg slide_idx} must be {.cls numeric}",
        "x" = "You provided {.cls {class(slide_idx)[1]}} instead.")
      , call = NULL)
  }
  n <- length(x) # no of slides
  check <- slide_idx %in% seq_len(n)
  if (!check) {
    cli::cli_abort(
      c("Slide index {.val {slide_idx}} is out of range.",
        "x" = "Presentation has {cli::no(n)} slide{?s}."
      ), call = NULL
    )
  }
}
