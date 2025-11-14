# PH_WITH -------

#' @export
#' @title Add objects on the current slide
#' @description add an object into a new shape in the current slide. This
#' function is able to add all supported outputs to a presentation. See
#' section **Methods (by class)** to see supported outputs.
#' @param x an rpptx object
#' @param value object to add as a new shape. Supported objects
#' are vectors, data.frame, graphics, block of formatted paragraphs,
#' unordered list of formatted paragraphs,
#' pretty tables with package flextable, editable graphics with
#' package rvg, 'Microsoft' charts with package mschart.
#' @param location a placeholder location object or a location short form. It will be used
#' to specify the location of the new shape. This location can be defined with a call to one
#' of the `ph_location_*` functions (see section `"see also"`). In `ph_with()`, several location
#' short forms can be used, as listed in section `"Short forms"`.
#' @param ... further arguments passed to or from other methods. When
#' adding a `ggplot` object or `plot_instr`, these arguments will be used
#' by the png function.
#'
#' @section Short forms:
#' The `location` argument of `ph_with()` either expects a location object as returned by the
#' `ph_location_*` functions or a corresponding location *short form* (string or numeric):
#'
#' | **Location function**                       | **Short form**             | **Description**                                      |
#' |---------------------------------------------|----------------------------|------------------------------------------------------|
#' | `ph_location_left()`                        | `"left"`                   | Keyword string                                       |
#' | `ph_location_right()`                       | `"right"`                  | Keyword string                                       |
#' | `ph_location_fullsize()`                    | `"fullsize"`               | Keyword string                                       |
#' | `ph_location_type("body", 1)`               | `"body [1]"`               | String: type + index in brackets (`1` if omitted)    |
#' | `ph_location_label("my_label")`             | `"my_label"`               | Any string not matching a keyword or type            |
#' | `ph_location_id(1)`                         | `1`                        | Length 1 integer                                     |
#' | `ph_location(0, 0, 4, 5)`                   | `c(0,0,4,5)`               | Length 4 numeric, optionally named, `c(top=0, left=0, ...)` |
#'
#' @example inst/examples/example_ph_with.R
#' @seealso Specify placeholder locations with [ph_location_type], [ph_location],
#' [ph_location_label], [ph_location_left], [ph_location_right],
#' [ph_location_fullsize], [ph_location_template]. [phs_with] is a sibling of
#' `ph_with` that fills multiple placeholders at once. Use [add_slide] to add new slides.
#' @section Illustrations:
#'
#' \if{html}{\figure{ph_with_doc_1.png}{options: style="width:80\%;"}}
ph_with <- function(x, value, location, ...) {
  location <- resolve_location(location)
  .ph_with(x, value, location, ...)
}


.ph_with <- function(x, value, location, ...) {
  UseMethod("ph_with", value)
}


#' @export
#' @describeIn ph_with add a character vector to a new shape on the
#' current slide, values will be added as paragraphs.
ph_with.character <- function(x, value, location, ...) {
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)
  new_ph <- shape_properties_tags(
    left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln, geom = location$geom
  )

  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscapeCopy(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0(
    psp_ns_yes, new_ph,
    "<p:txBody><a:bodyPr/><a:lstStyle/>",
    pars, "</p:txBody></p:sp>"
  )

  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  x
}

#' @export
#' @param format_fun format function for non character vectors
#' @describeIn ph_with add a numeric vector to a new shape on the
#' current slide, values will be be first formatted then
#' added as paragraphs.
ph_with.numeric <- function(x, value, location, format_fun = format, ...) {
  slide <- x$slide$get_slide(x$cursor)
  value <- format_fun(value, ...)
  location <- fortify_location(location, doc = x)

  new_ph <- shape_properties_tags(
    left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln, geom = location$geom
  )

  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscapeCopy(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0(
    psp_ns_yes, new_ph,
    "<p:txBody><a:bodyPr/><a:lstStyle/>",
    pars, "</p:txBody></p:sp>"
  )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  x
}

#' @export
#' @describeIn ph_with add a factor vector to a new shape on the
#' current slide, values will be be converted as character and then
#' added as paragraphs.
ph_with.factor <- function(x, value, location, ...) {
  slide <- x$slide$get_slide(x$cursor)
  value <- as.character(value)
  location <- fortify_location(location, doc = x)

  new_ph <- shape_properties_tags(
    left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln, geom = location$geom
  )

  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscapeCopy(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- paste0(
    psp_ns_yes, new_ph,
    "<p:txBody><a:bodyPr/><a:lstStyle/>",
    pars, "</p:txBody></p:sp>"
  )
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  x
}
#' @export
#' @rdname ph_with
ph_with.logical <- ph_with.numeric


#' @export
#' @param date_format A format string for dates (default `"%Y-%m-%d"`).
#' See `format` arg in [strftime()] for details.
#' Set a global default via `options(officer.date_format = ...)`.
#' @describeIn ph_with add a `Date` object vector to a new shape on the
#' current slide, values will be be first converted to character.
ph_with.Date <- function(x, value, location, date_format = NULL, ...) {
  opt <- options()$officer.date_format
  if (!is.null(opt) && is.na(opt)) { # catch NA
    opt <- NULL
  }
  format <- date_format %||% opt %||% "%Y-%m-%d" # fallback to format.Date() default
  value_str <- format(value, format = format)
  ph_with(x, value = value_str, location, ...)
}


#' @export
#' @param level_list The list of levels for hierarchy structure as integer values.
#' If used the object is formated as an unordered list. If 1 and 2,
#' item 1 level will be 1, item 2 level will be 2.
#' @describeIn ph_with add a [block_list()] made
#' of [fpar()] to a new shape on the current slide.
ph_with.block_list <- function(x, value, location, level_list = integer(0), ...) {
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)

  pars <- sapply(value, to_pml)

  if (length(level_list) > 0) {
    pars <- gsub("<a:buNone/>", "", pars, fixed = TRUE)

    level_values <- rep(1L, length(pars))
    level_values[] <- level_list
    level_values <- level_values - 1L

    lvl <- sprintf(" lvl=\"%.0f\"", level_values)
    lvl <- paste0("<a:pPr", ifelse(level_values > 0, lvl, ""), "/>")

    pars <- mapply(
      function(par, lvl) {
        paste0(par[1], lvl, par[2])
      }, strsplit(pars, split = "<a:pPr(.*)</a:pPr>"), lvl,
      SIMPLIFY = FALSE
    )
    pars <- unlist(pars)
  }

  pars <- paste0(pars, collapse = "")

  new_ph <- shape_properties_tags(
    left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln, geom = location$geom
  )

  xml_elt <- paste0(
    psp_ns_yes, new_ph,
    "<p:txBody><a:bodyPr/><a:lstStyle/>",
    pars, "</p:txBody></p:sp>"
  )

  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  x
}


#' @export
#' @describeIn ph_with add a [unordered_list()] made
#' of [fpar()] to a new shape on the current slide.
ph_with.unordered_list <- function(x, value, location, ...) {
  slide <- x$slide$get_slide(x$cursor)
  location <- fortify_location(location, doc = x)

  p <- to_pml(value)

  new_ph <- shape_properties_tags(
    left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln, geom = location$geom
  )

  xml_elt <- paste0(
    psp_ns_yes, new_ph,
    "<p:txBody><a:bodyPr/><a:lstStyle/>", p, "</p:txBody></p:sp>"
  )
  node <- as_xml_document(xml_elt)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  x
}


#' @export
#' @param header display header if TRUE
#' @param tcf conditional formatting settings defined by [table_conditional_formatting()]
#' @param alignment alignment for each columns, 'l' for left, 'r' for right
#' and 'c' for center. Default to NULL.
#' @describeIn ph_with add a data.frame to a new shape on the current slide with
#' function [block_table()]. Use package 'flextable' instead for more
#' advanced formattings.
ph_with.data.frame <- function(x, value, location, header = TRUE,
                               tcf = table_conditional_formatting(),
                               alignment = NULL,
                               ...) {
  location <- fortify_location(location, doc = x)

  slide <- x$slide$get_slide(x$cursor)
  style_id <- x$table_styles$def[1]

  pt <- prop_table(
    style = style_id, layout = table_layout(),
    width = table_width(),
    tcf = tcf
  )

  bt <- block_table(x = value, header = header, properties = pt, alignment = alignment)

  xml_elt <- to_pml(
    bt,
    left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln, geom = location$geom
  )

  value <- as_xml_document(xml_elt)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)
  x
}


#' @export
#' @describeIn ph_with add a ggplot object to a new shape on the
#' current slide. Use package 'rvg' for more advanced graphical features.
#' @param res resolution of the png image in ppi
#' @param alt_text Alt-text for screen-readers. Defaults to `""`. If `""` or `NULL`
#'    an alt text added with `ggplot2::labs(alt = ...)` will be used if any.
#' @param scale Multiplicative scaling factor, same as in ggsave
ph_with.gg <- function(x, value, location, res = 300, alt_text = "", scale = 1, ...) {
  location_ <- fortify_location(location, doc = x)
  slide <- x$slide$get_slide(x$cursor)
  if (!requireNamespace("ggplot2")) {
    stop("package ggplot2 is required to use this function")
  }

  slide <- x$slide$get_slide(x$cursor)
  width <- location_$width
  height <- location_$height

  stopifnot(inherits(value, "gg"))
  file <- tempfile(fileext = ".png")
  agg_png(filename = file, width = width, height = height, units = "in", res = res, scaling = scale, background = "transparent", ...)
  print(value)
  dev.off()
  on.exit(unlink(file))

  if (is.null(alt_text) || alt_text == "") {
    alt_text <- ggplot2::get_alt_text(value)
    if (is.null(alt_text)) alt_text <- ""
  }

  ext_img <- external_img(file, width = width, height = height, alt = alt_text)

  ph_with(x, ext_img, location = location)
}

#' @export
#' @describeIn ph_with add an R plot to a new shape on the
#' current slide. Use package 'rvg' for more advanced graphical features.
ph_with.plot_instr <- function(x, value, location, res = 300, ...) {
  location_ <- fortify_location(location, doc = x)
  slide <- x$slide$get_slide(x$cursor)
  slide <- x$slide$get_slide(x$cursor)
  width <- location_$width
  height <- location_$height

  file <- tempfile(fileext = ".png")

  dirname <- tempfile()
  dir.create(dirname)
  filename <- paste(dirname, "/plot%03d.png", sep = "")
  agg_png(filename = filename, width = width, height = height, units = "in", res = res, scaling = 1, background = "transparent", ...)

  tryCatch(
    {
      eval(value$code)
    },
    finally = {
      dev.off()
    }
  )
  file <- list.files(dirname, full.names = TRUE)
  on.exit(unlink(dirname, recursive = TRUE, force = TRUE))

  if (length(file) > 1) {
    stop(length(file), " files have been produced. Multiple plot are not supported")
  }

  ext_img <- external_img(file, width = width, height = height)
  ph_with(x, ext_img, location = location)
}


#' @export
#' @param use_loc_size if set to FALSE, external_img width and height will
#' be used.
#' @describeIn ph_with add a [external_img()] to a new shape
#' on the current slide.
#'
#' When value is a external_img object, image will be copied
#' into the PowerPoint presentation. The width and height
#' specified in call to [external_img()] will be
#' ignored, their values will be those of the location,
#' unless use_loc_size is set to FALSE.
ph_with.external_img <- function(x, value, location, use_loc_size = TRUE, ...) {
  location <- fortify_location(location, doc = x)

  slide <- x$slide$get_slide(x$cursor)

  if (!use_loc_size) {
    location$width <- attr(value, "dims")$width
    location$height <- attr(value, "dims")$height
  }
  width <- location$width
  height <- location$height


  xml_str <- to_pml(
    x = value,
    left = location$left, top = location$top,
    width = width, height = height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln
  )

  value <- as_xml_document(xml_str)
  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)
  x
}




#' @export
#' @describeIn ph_with add an [fpar()] to a new shape
#' on the current slide as a single paragraph in a [block_list()].
ph_with.fpar <- function(x, value, location, ...) {
  ph_with.block_list(x, value = block_list(value), location = location)

  x
}

#' @export
#' @describeIn ph_with add an [empty_content()] to a new shape
#' on the current slide.
ph_with.empty_content <- function(x, value, location, ...) {
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)
  new_ph <- shape_properties_tags(
    left = location$left, top = location$top,
    width = location$width, height = location$height,
    label = location$ph_label, ph = location$ph,
    rot = location$rotation, bg = location$bg,
    ln = location$ln, geom = location$geom
  )

  if (is.na(location$fld_id)) {
    xml_elt <- paste0(psp_ns_yes, new_ph, "</p:sp>")
  } else {
    pars <- paste0("<a:p><a:fld id=\"", location$fld_id, "\" type = \"", location$fld_type, "\"><a:rPr/><a:t>", x$cursor, "</a:t></a:fld></a:p>", collapse = "")
    xml_elt <- sprintf(paste0(
      psp_ns_yes, new_ph,
      "<p:txBody><a:bodyPr/><a:lstStyle/>",
      pars, "</p:txBody></p:sp>"
    ))
  }
  node <- as_xml_document(xml_elt)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), node)
  x
}



xml_to_slide <- function(slide, location, value, package_dir) {
  node <- xml_find_first(value, as_xpath_content_sel("//"))
  if (xml_name(node) == "grpSp") {
    node_sppr <- xml_child(node, "p:grpSpPr")
    node_xfrm <- xml_child(node, "p:grpSpPr/a:xfrm")
    node_name <- xml_child(node, "p:nvGrpSpPr/p:cNvPr")
  } else if (xml_name(node) == "graphicFrame") {
    node_xfrm <- xml_child(node, "p:xfrm")
    node_name <- xml_child(node, "p:nvGraphicFramePr/p:cNvPr")
    node_sppr <- xml_missing()
  } else if (xml_name(node) == "pic") {
    node_sppr <- xml_child(node, "p:spPr")
    node_xfrm <- xml_child(node, "p:spPr/a:xfrm")
    node_name <- xml_child(node, "p:nvPicPr/p:cNvPr")
  } else if (xml_name(node) == "sp") {
    node_sppr <- xml_child(node, "p:spPr")
    node_xfrm <- xml_child(node, "p:spPr/a:xfrm")
    node_name <- xml_child(node, "p:nvSpPr/p:cNvPr")
  } else {
    node_xfrm <- xml_missing()
    node_name <- xml_missing()
    node_sppr <- xml_missing()
  }


  if (!inherits(node_name, "xml_missing")) {
    xml_attr(node_name, "name") <- location$ph_label
  }

  if (!inherits(node_xfrm, "xml_missing")) {
    off <- xml_child(node_xfrm, "a:off")
    ext <- xml_child(node_xfrm, "a:ext")
    chOff <- xml_child(node_xfrm, "a:chOff")
    chExt <- xml_child(node_xfrm, "a:chExt")
    xml_attr(off, "x") <- sprintf("%.0f", location$left * 914400)
    xml_attr(off, "y") <- sprintf("%.0f", location$top * 914400)
    xml_attr(ext, "cx") <- sprintf("%.0f", location$width * 914400)
    xml_attr(ext, "cy") <- sprintf("%.0f", location$height * 914400)
    xml_attr(chOff, "x") <- sprintf("%.0f", location$left * 914400)
    xml_attr(chOff, "y") <- sprintf("%.0f", location$top * 914400)
    xml_attr(chExt, "cx") <- sprintf("%.0f", location$width * 914400)
    xml_attr(chExt, "cy") <- sprintf("%.0f", location$height * 914400)

    # add location$rotation to cNvPr:name
    if (!is.null(location$rotation) &&
      is.finite(location$rotation) &&
      is.numeric(location$rotation)) {
      xml_attr(node_xfrm, "rot") <- sprintf("%.0f", -location$rotation * 60000)
    }
  }

  if (!inherits(node_sppr, "xml_missing")) {
    # add location$bg to SpPr
    if (!is.null(location$bg)) {
      bg_str <- solid_fill_pml(location$bg)
      xml_add_child(node_sppr, as_xml_document(bg_str))
    }
  }
  value
}

#' @export
#' @describeIn ph_with add an xml_document object to a new shape on the
#' current slide. This function is to be used to add custom openxml code.
ph_with.xml_document <- function(x, value, location, ...) {
  slide <- x$slide$get_slide(x$cursor)

  location <- fortify_location(location, doc = x)

  xml_to_slide(slide, location, value, x$package_dir)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), value)
  x
}


# PHS_WITH -------

#' @title Fill multiple placeholders using key value syntax
#' @description A sibling of [officer::ph_with] that fills mutiple placeholders at once. Placeholder locations are
#'   specfied using the short form syntax. The location and corresponding object are passed as key value pairs
#'   (`phs_with("short form location" = object)`). Under the hood, [officer::ph_with] is called for each pair. Note
#'   that `phs_with` does not cover all options from the `ph_location_*` family and is also less customization. It is a
#'   covenience wrapper for the most common use cases. The implemented short forms are listed in section
#'   `"Short forms"`.
#' @param x A `rpptx` object.
#' @param ... Key-value pairs of the form `"short form location" = object`. If the short form is an integer or a string
#'   with blanks, you must wrap it in quotes or backticks.
#' @param .dots List of key-value pairs `"short form location" = object`. Alternative to `...`.
#' @param .slide_idx Numeric indexes of slides to process. `NULL` (default) processes the current slide only. Use
#'   keyword `all` for all slides.
#' @section Short forms: The following short forms are implemented and can be used as the parameter in the function
#'   call. The corresponding function from the `ph_location_*` family (called under the hood) is displayed on the
#'   right.
#'
#' | **Short form** | **Description**                                   | **Location function**           |
#' |----------------|---------------------------------------------------|---------------------------------|
#' | `"left"`       | Keyword string                                    | `ph_location_left()`            |
#' | `"right"`      | Keyword string                                    | `ph_location_right()`           |
#' | `"fullsize"`   | Keyword string                                    | `ph_location_fullsize()`        |
#' | `"body [1]"`   | String: type + index in brackets (`1` if omitted) | `ph_location_type("body", 1)`   |
#' | `"my_label"`   | Any string not matching a keyword or type         | `ph_location_label("my_label")` |
#' | `1`            | Length 1 integer                                  | `ph_location_id(1)`             |
#'
#' @export
#' @seealso [ph_with()], [add_slide()]
#' @example inst/examples/example_phs_with.R
phs_with <- function(x, ..., .dots = NULL, .slide_idx = NULL) {
  dots_list <- list(...)
  if (length(dots_list) > 0 && !is_named(dots_list)) {
    cli::cli_abort(
      c("Missing key in {.arg ...}",
        "x" = "{.arg ...} requires a key-value syntax, with a ph short-form location as the key",
        "i" = "Example: {.code phs_with(x, title = 'My title', 'body[1]' = 'My body')}"
      )
    )
  }
  if (length(.dots) > 0 && !is_named(.dots)) {
    cli::cli_abort(
      c("Missing names in {.arg .dots}",
        "x" = "{.arg .dots} must be a named list, with ph short-form locations as names",
        "i" = "Example: {.code phs_with(x, .dots = list(title = 'My title', 'body[1]' = 'My body'))}"
      )
    )
  }
  dots <- utils::modifyList(dots_list, .dots %||% list())
  if (length(dots) == 0) {
    return(x)
  }
  .slide_idx <- .slide_idx %||% x$cursor # default is current slide
  if (is.character(.slide_idx) && .slide_idx == "all") {
    .slide_idx <- seq_len(length(x))
  }
  stop_if_not_in_slide_range(x, .slide_idx)

  loc_strings <- as.list(names(dots))
  ii <- grepl("^\\d+$", loc_strings) # find integer short-forms
  loc_strings[ii] <- as.integer(loc_strings[ii])
  locations <- lapply(loc_strings, resolve_location)

  .old_cursor <- x$cursor
  for (slide_idx in .slide_idx) {
    x$cursor <- slide_idx # pw_with always uses the current slide
    for (i in seq_along(dots)) {
      x <- ph_with(x, dots[[i]], locations[[i]])
    }
  }
  x$cursor <- .old_cursor
  x
}
