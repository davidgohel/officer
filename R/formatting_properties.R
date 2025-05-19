# check properties helpers ----

check_spread_integer <- function(obj, value, dest) {
  varname <- as.character(substitute(value))
  if (is.numeric(value) && length(value) == 1 && value >= 0) {
    for (i in dest) {
      obj[[i]] <- as.integer(value)
    }
  } else if (length(value) == 1 && is.na(value)) {
    for (i in dest) {
      obj[[varname]] <- value
    }
  } else {
    stop(varname, " must be a positive integer scalar.", call. = FALSE)
  }
  obj
}

check_set_color <- function(obj, value) {
  varname <- as.character(substitute(value))
  if (!is.color(value) && !is.na(value)) {
    stop(varname, " must be a valid color.", call. = FALSE)
  } else {
    obj[[varname]] <- value
  }
  obj
}


check_set_border <- function(obj, value) {
  varname <- as.character(substitute(value))
  if (!inherits(value, "fp_border")) {
    stop(varname, " must be a fp_border object.", call. = FALSE)
  } else {
    obj[[varname]] <- value
  }
  obj
}

check_spread_border <- function(obj, value, dest) {
  varname <- as.character(substitute(value))
  if (!inherits(value, "fp_border")) {
    stop(varname, " must be a fp_border object.", call. = FALSE)
  }
  for (i in dest) {
    obj[[i]] <- value
  }
  obj
}

check_set_integer <- function(obj, value) {
  varname <- as.character(substitute(value))
  if (is.na(value) || (is.numeric(value) && length(value) == 1 && value >= 0)) {
    obj[[varname]] <- as.integer(value)
  } else {
    stop(varname, " must be a positive integer scalar.", call. = FALSE)
  }
  obj
}

check_set_numeric <- function(obj, value) {
  varname <- as.character(substitute(value))
  if (is.na(value) || (is.numeric(value) && length(value) == 1 && value >= 0)) {
    obj[[varname]] <- as.double(value)
  } else {
    stop(varname, " must be a positive numeric scalar.", call. = FALSE)
  }
  obj
}


check_set_bool <- function(obj, value) {
  varname <- as.character(substitute(value))
  if (is.na(value) || (is.logical(value) && length(value) == 1)) {
    obj[[varname]] <- value
  } else {
    stop(varname, " must be a boolean", call. = FALSE)
  }
  obj
}
check_set_chr <- function(obj, value) {
  varname <- as.character(substitute(value))
  if (is.na(value) || (is.character(value) && length(value) == 1)) {
    obj[[varname]] <- value
  } else {
    stop(varname, " must be a string", call. = FALSE)
  }
  obj
}

check_set_choice <- function(obj, value, choices) {
  varname <- as.character(substitute(value))

  if (is.na(value)) {
    obj[[varname]] <- value
  } else {
    if (is.character(value) && length(value) == 1) {
      if (!value %in% choices) {
        stop(
          varname,
          " must be one of ",
          paste(shQuote(choices), collapse = ", "),
          call. = FALSE
        )
      }
      obj[[varname]] <- value
    } else {
      stop(varname, " must be a character scalar.", call. = FALSE)
    }
  }

  obj
}

check_set_class <- function(obj, value, cl) {
  varname <- as.character(substitute(value))
  cl_str <- sprintf(" must be a %s object.", cl)
  if (!inherits(value, cl)) {
    stop(varname, cl_str, call. = FALSE)
  } else {
    obj[[varname]] <- value
  }
  obj
}

default_rpr <- data.frame(
  stringsAsFactors = FALSE,
  font.size = NA_integer_,
  bold = as.logical(NA),
  italic = as.logical(NA),
  underlined = as.logical(NA),
  color = NA_character_,
  font.family = NA_character_,
  bold.cs = as.logical(NA),
  font.size.cs = NA_integer_,
  vertical.align = NA_character_,
  shading.color = NA_character_,
  hansi.family = NA_character_,
  eastasia.family = NA_character_,
  cs.family = NA_character_,
  lang.val = NA_character_,
  lang.eastasia = NA_character_,
  lang.bidi = NA_character_
)

is_positive_numeric <- function(x) {
  varname <- as.character(substitute(x))
  test <- is.na(x) || (is.numeric(x) && length(x) == 1 && x >= 0)
  if (!test) {
    stop(varname, " must be a positive numeric scalar.", call. = FALSE)
  }
  test
}
is_bool <- function(x) {
  varname <- as.character(substitute(x))
  test <- is.na(x) || (is.logical(x) && length(x) == 1)
  if (!test) {
    stop(varname, " must be a boolean.", call. = FALSE)
  }
  test
}
is_color <- function(x) {
  varname <- as.character(substitute(x))
  test <- is.na(x) || is.color(x)
  if (!test) {
    stop(varname, " must be a valid color.", call. = FALSE)
  }
  test
}
is_character <- function(x) {
  varname <- as.character(substitute(x))
  test <- is.na(x) || (is.character(x) && length(x) == 1)
  if (!test) {
    stop(varname, " must be a string", call. = FALSE)
  }
  test
}
# fp_text ----
#' @title Text formatting properties
#' @description Create an `fp_text` object that describes
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
#' Expected value is one of the following : default `'baseline'`
#' or `'subscript'` or `'superscript'`
#' @param shading.color shading color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @return a `fp_text` object
#' @examples
#' fp_text()
#' fp_text(color = "red")
#' fp_text(bold = TRUE, shading.color = "yellow")
#' @family functions for defining formatting properties
#' @seealso [ftext()], [fpar()]
#' @export
fp_text <- function(
  color = "black",
  font.size = 10,
  bold = FALSE,
  italic = FALSE,
  underlined = FALSE,
  font.family = "Arial",
  cs.family = NULL,
  eastasia.family = NULL,
  hansi.family = NULL,
  vertical.align = "baseline",
  shading.color = "transparent"
) {
  out <- default_rpr
  if (is_positive_numeric(font.size)) {
    out$font.size <- font.size
    out$font.size.cs <- font.size
  }
  if (is_bool(bold)) {
    out$bold <- bold
    out$bold.cs <- bold
  }
  if (is_bool(italic)) {
    out$italic <- italic
  }
  if (is_bool(underlined)) {
    out$underlined <- underlined
  }
  if (is_color(color)) {
    out$color <- color
  }
  if (is_character(font.family)) {
    out$font.family <- font.family
  }

  if (is.null(cs.family)) cs.family <- font.family
  if (is.null(eastasia.family)) eastasia.family <- font.family
  if (is.null(hansi.family)) hansi.family <- font.family
  if (is_character(cs.family)) {
    out$cs.family <- cs.family
  }
  if (is_character(eastasia.family)) {
    out$eastasia.family <- eastasia.family
  }
  if (is_character(hansi.family)) {
    out$hansi.family <- hansi.family
  }

  out <- check_set_choice(
    obj = out,
    value = vertical.align,
    choices = c("subscript", "superscript", "baseline")
  )
  out <- check_set_color(out, shading.color)

  class(out) <- "fp_text"

  out
}


#' @rdname fp_text
#' @description Function `fp_text_lite()` is generating properties
#' with only entries for the parameters users provided. The
#' undefined properties will inherit from the default settings.
#' @export
fp_text_lite <- function(
  color = NA,
  font.size = NA,
  font.family = NA,
  cs.family = NA,
  eastasia.family = NA,
  hansi.family = NA,
  bold = NA,
  italic = NA,
  underlined = NA,
  vertical.align = "baseline",
  shading.color = NA
) {
  fp_text(
    color = color,
    font.size = font.size,
    bold = bold,
    italic = italic,
    underlined = underlined,
    font.family = font.family,
    cs.family = cs.family,
    eastasia.family = eastasia.family,
    hansi.family = hansi.family,
    vertical.align = vertical.align,
    shading.color = shading.color
  )
}

#' @rdname fp_text
#' @param format format type, wml for MS word, pml for
#' MS PowerPoint and html.
#' @param type output type - one of 'wml', 'pml', 'html', 'rtf'.
#' @export
format.fp_text <- function(x, type = "wml", ...) {
  stopifnot(length(type) == 1)
  stopifnot(type %in% c("wml", "pml", "html", "rtf"))

  if (type == "wml") {
    rpr_wml(x)
  } else if (type == "pml") {
    rpr_pml(x)
  } else if (type == "html") {
    rpr_css(x)
  } else if (type == "rtf") {
    rpr_rtf(x)
  } else {
    stop("unimplemented type")
  }
}

#' @export
to_wml.fp_text <- function(x, add_ns = FALSE, ...) {
  format(x, type = "wml")
}

#' @param x `fp_text` object
#' @examples
#' print(fp_text(color = "red", font.size = 12))
#' @rdname fp_text
#' @export
print.fp_text <- function(x, ...) {
  out <- data.frame(
    font.size = as.double(x$font.size),
    italic = x$italic,
    bold = x$bold,
    underlined = x$underlined,
    color = x$color,
    shading = x$shading.color,
    fontname = x$font.family,
    fontname_cs = x$cs.family,
    fontname_eastasia = x$eastasia.family,
    fontname.hansi = x$hansi.family,
    vertical_align = x$vertical.align,
    stringsAsFactors = FALSE
  )
  print(out)
  invisible()
}


#' @param object `fp_text` object to modify
#' @param ... further arguments - not used
#' @rdname fp_text
#' @export
update.fp_text <- function(
  object,
  color,
  font.size,
  bold,
  italic,
  underlined,
  font.family,
  cs.family,
  eastasia.family,
  hansi.family,
  vertical.align,
  shading.color,
  ...
) {
  if (!missing(font.size)) {
    object <- check_set_numeric(obj = object, font.size)
  }
  if (!missing(bold)) {
    object <- check_set_bool(obj = object, bold)
  }
  if (!missing(italic)) {
    object <- check_set_bool(obj = object, italic)
  }
  if (!missing(underlined)) {
    object <- check_set_bool(obj = object, underlined)
  }
  if (!missing(color)) {
    object <- check_set_color(object, color)
  }
  if (!missing(font.family)) {
    object <- check_set_chr(object, font.family)
  }
  if (!missing(cs.family)) {
    object <- check_set_chr(object, cs.family)
  }
  if (!missing(eastasia.family)) {
    object <- check_set_chr(object, eastasia.family)
  }
  if (!missing(hansi.family)) {
    object <- check_set_chr(object, hansi.family)
  }
  if (!missing(vertical.align)) {
    object <- check_set_choice(
      obj = object,
      value = vertical.align,
      choices = c("subscript", "superscript", "baseline")
    )
  }
  if (!missing(shading.color)) {
    object <- check_set_color(object, shading.color)
  }

  object
}

# fp_border ----
border_styles <- c(
  "nil",
  "none",
  "solid",
  "single",
  "thick",
  "double",
  "dotted",
  "dashed",
  "dotDash",
  "dotDotDash",
  "triple",
  "thinThickSmallGap",
  "thickThinSmallGap",
  "thinThickThinSmallGap",
  "thinThickMediumGap",
  "thickThinMediumGap",
  "thinThickThinMediumGap",
  "thinThickLargeGap",
  "thickThinLargeGap",
  "thinThickThinLargeGap",
  "wave",
  "doubleWave",
  "dashSmallGap",
  "dashDotStroked",
  "threeDEmboss",
  "ridge",
  "threeDEngrave",
  "groove",
  "outset",
  "inset"
)

#' @title Border properties object
#'
#' @description create a border properties object.
#'
#' @param color border color - single character value (e.g. "#000000" or "black")
#' @param style border style - single character value : See Details for supported
#'  border styles.
#' @param width border width - an integer value : 0>= value
#' @details For Word output the following border styles are supported:
#'
#' * "none" or "nil" - No Border
#' * "solid" or "single" - Single Line Border
#' * "thick" - Single Line Border
#' * "double" - Double Line Border
#' * "dotted" - Dotted Line Border
#' * "dashed" - Dashed Line Border
#' * "dotDash" - Dot Dash Line Border
#' * "dotDotDash" - Dot Dot Dash Line Border
#' * "triple" - Triple Line Border
#' * "thinThickSmallGap" - Thin, Thick Line Border
#' * "thickThinSmallGap" - Thick, Thin Line Border
#' * "thinThickThinSmallGap" - Thin, Thick, Thin Line Border
#' * "thinThickMediumGap" - Thin, Thick Line Border
#' * "thickThinMediumGap" - Thick, Thin Line Border
#' * "thinThickThinMediumGap" - Thin, Thick, Thin Line Border
#' * "thinThickLargeGap" - Thin, Thick Line Border
#' * "thickThinLargeGap" - Thick, Thin Line Border
#' * "thinThickThinLargeGap" - Thin, Thick, Thin Line Border
#' * "wave" - Wavy Line Border
#' * "doubleWave" - Double Wave Line Border
#' * "dashSmallGap" - Dashed Line Border
#' * "dashDotStroked" - Dash Dot Strokes Line Border
#' * "threeDEmboss" or "ridge" - 3D Embossed Line Border
#' * "threeDEngrave" or "groove" - 3D Engraved Line Border
#' * "outset" - Outset Line Border
#' * "inset" - Inset Line Border
#'
#' For HTML output only a limited amount of border styles are supported:
#'
#' * "none" or "nil" - No Border
#' * "solid" or "single" - Single Line Border
#' * "double" - Double Line Border
#' * "dotted" - Dotted Line Border
#' * "dashed" - Dashed Line Border
#' * "threeDEmboss" or "ridge" - 3D Embossed Line Border
#' * "threeDEngrave" or "groove" - 3D Engraved Line Border
#' * "outset" - Outset Line Border
#' * "inset" - Inset Line Border
#'
#' Non-supported Word border styles will default to "solid".
#' @examples
#' fp_border()
#' fp_border(color = "orange", style = "solid", width = 1)
#' fp_border(color = "gray", style = "dotted", width = 1)
#' @export
#' @family functions for defining formatting properties
fp_border <- function(color = "black", style = "solid", width = 1) {
  out <- list()
  out <- check_set_numeric(obj = out, width)
  out <- check_set_color(out, color)
  out <- check_set_choice(
    obj = out,
    style,
    choices = border_styles
  )

  class(out) <- "fp_border"
  out
}

#' @param object fp_border object
#' @param ... further arguments - not used
#' @rdname fp_border
#' @examples
#'
#' # modify object ------
#' border <- fp_border()
#' update(border, style = "dotted", width = 3)
#' @export
update.fp_border <- function(object, color, style, width, ...) {
  if (!missing(color)) {
    object <- check_set_color(object, color)
  }

  if (!missing(width)) {
    object <- check_set_integer(obj = object, width)
  }

  if (!missing(style)) {
    object <- check_set_choice(obj = object, style, choices = border_styles)
  }

  object
}

#' @export
print.fp_border <- function(x, ...) {
  msg <- paste0(
    "line: color: ",
    x$color,
    ", width: ",
    x$width,
    ", style: ",
    x$style,
    "\n"
  )
  cat(msg)
  invisible()
}

# tabs ----
#' @export
#' @title Tabulation mark properties object
#'
#' @description create a tabulation mark properties setting object for Word
#' or RTF. Results can be used as arguments of [fp_tabs()].
#'
#' Once tabulation marks settings are defined, tabulation marks can
#' be added with [run_tab()] inside a call to [fpar()] or
#' with `\t` within 'flextable' content.
#'
#' @param pos Specifies the position of the tab stop (in inches).
#' @param style style of the tab. Possible values are:
#' "decimal", "left", "right" or "center".
#' @examples
#' fp_tab(pos = 0.4, style = "decimal")
#' fp_tab(pos = 1, style = "right")
#' @family functions for defining formatting properties
fp_tab <- function(pos, style = "decimal") {
  out <- list()
  out <- check_set_numeric(obj = out, pos)
  out <- check_set_choice(
    obj = out,
    style,
    choices = c(
      "decimal",
      "left",
      "right",
      "center"
    )
  )

  class(out) <- "fp_tab"
  out
}

#' @export
to_wml.fp_tab <- function(x, add_ns = FALSE, ...) {
  sprintf(
    "<w:tab w:val=\"%s\" w:leader=\"none\" w:pos=\"%.0f\"/>",
    x$style,
    inch_to_tweep(x$pos)
  )
}
#' @export
as.character.fp_tab <- function(x, ...) {
  paste(x$style, x$pos, sep = "_")
}


#' @export
#' @title Tabs properties object
#'
#' @description create a set of tabulation mark properties object for Word or RTF.
#' Results can be used as arguments `tabs` of [fp_par()] and will only have
#' effects in Word or RTF outputs.
#'
#' Once a set of tabulation marks settings is defined, tabulation marks can
#' be added with [run_tab()] inside a call to [fpar()] or
#' with `\t` within 'flextable' content.
#' @param ... [fp_tab] objects
#' @examples
#' z <- fp_tabs(
#'   fp_tab(pos = 0.4, style = "decimal"),
#'   fp_tab(pos = 1, style = "decimal")
#' )
#' fpar(
#'   run_tab(), ftext("88."),
#'   run_tab(), ftext("987.45"),
#'   fp_p = fp_par(
#'     tabs = z
#'   )
#' )
#' @family functions for defining formatting properties
fp_tabs <- function(...) {
  out <- list(...)
  if (length(out) > 0) {
    are_fp_tab <- sapply(out, inherits, "fp_tab")
    if (!all(are_fp_tab)) {
      stop("Function `fp_tabs()` only accept `fp_tab` objects as arguments.")
    }
  }

  class(out) <- "fp_tabs"
  out
}

#' @export
to_wml.fp_tabs <- function(x, add_ns = FALSE, ...) {
  if (length(x) > 0) {
    paste0("<w:tabs>", do.call(paste0, lapply(x, to_wml)), "</w:tabs>")
  } else {
    ""
  }
}

#' @export
as.character.fp_tabs <- function(x, ...) {
  if (length(x) > 0) {
    z <- lapply(x, as.character)
    z$sep <- "&"
    do.call(paste, z)
  } else {
    NA_character_
  }
}

# fp_par -----
#' @title Paragraph formatting properties
#'
#' @description Create a `fp_par` object that describes
#' paragraph formatting properties.
#'
#' @param text.align text alignment - a single character value, expected value
#' is one of 'left', 'right', 'center', 'justify'.
#' @param padding.bottom,padding.top,padding.left,padding.right paragraph paddings - 0 or positive integer value.
#' @param padding paragraph paddings - 0 or positive integer value. Argument `padding` overwrites
#' arguments `padding.bottom`, `padding.top`, `padding.left`, `padding.right`.
#' @param line_spacing line spacing, 1 is single line spacing, 2 is double line spacing.
#' @param border shortcut for all borders.
#' @param border.bottom,border.left,border.top,border.right [fp_border()] for
#' borders. overwrite other border properties.
#' @param shading.color shading color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @param keep_with_next a scalar logical. Specifies that the paragraph (or at least part of it) should be rendered
#' on the same page as the next paragraph when possible.
#' @param tabs NULL (default) for no tabulation marks setting
#' or an object returned by [fp_tabs()]. Note this can only have effect with Word
#' or RTF outputs.
#' @param word_style Word paragraph style name
#' @return a `fp_par` object
#' @examples
#' fp_par(text.align = "center", padding = 5)
#' @export
#' @family functions for defining formatting properties
#' @seealso [fpar]
fp_par <- function(
  text.align = "left",
  padding = 0,
  line_spacing = 1,
  border = fp_border(width = 0),
  padding.bottom,
  padding.top,
  padding.left,
  padding.right,
  border.bottom,
  border.left,
  border.top,
  border.right,
  shading.color = "transparent",
  keep_with_next = FALSE,
  tabs = NULL,
  word_style = "Normal"
) {
  out <- list()

  out <- check_set_color(out, shading.color)
  out <- check_set_choice(
    obj = out,
    value = text.align,
    choices = c("left", "right", "center", "justify")
  )
  # padding checking
  out <- check_spread_integer(
    out,
    padding,
    c(
      "padding.bottom",
      "padding.top",
      "padding.left",
      "padding.right"
    )
  )
  if (!missing(padding.bottom)) {
    out <- check_set_numeric(obj = out, padding.bottom)
  }
  if (!missing(padding.left)) {
    out <- check_set_numeric(obj = out, padding.left)
  }
  if (!missing(padding.top)) {
    out <- check_set_numeric(obj = out, padding.top)
  }
  if (!missing(padding.right)) {
    out <- check_set_numeric(obj = out, padding.right)
  }

  out <- check_set_numeric(obj = out, line_spacing)

  # border checking
  if (!is.null(border) && !isFALSE(border)) {
    out <- check_spread_border(
      obj = out,
      border,
      dest = c(
        "border.bottom",
        "border.top",
        "border.left",
        "border.right"
      )
    )
  }


  if (!missing(border.top) && !isFALSE(border.top)) {
    out <- check_set_border(obj = out, border.top)
  }
  if (!missing(border.bottom) && !isFALSE(border.bottom)) {
    out <- check_set_border(obj = out, border.bottom)
  }
  if (!missing(border.left) && !isFALSE(border.left)) {
    out <- check_set_border(obj = out, border.left)
  }
  if (!missing(border.right) && !isFALSE(border.right)) {
    out <- check_set_border(obj = out, border.right)
  }

  if (!is.null(tabs)) {
    out[["tabs"]] <- tabs
  }

  out <- check_set_chr(obj = out, word_style)

  out$keep_with_next <- keep_with_next
  class(out) <- "fp_par"

  out
}


#' @rdname fp_par
#' @description Function `fp_par_lite()` is generating properties
#' with only entries for the parameters users provided. The
#' undefined properties will inherit from the default settings.
#' @export
fp_par_lite <- function(
    text.align = NA,
    padding = NA,
    line_spacing = NA,
    border = FALSE,
    padding.bottom = NA,
    padding.top = NA,
    padding.left = NA,
    padding.right = NA,
    border.bottom = FALSE,
    border.left = FALSE,
    border.top = FALSE,
    border.right = FALSE,
    shading.color = NA,
    keep_with_next = NA,
    tabs = FALSE,
    word_style = NA
) {
  fp_par(
    text.align = text.align,
    padding = padding,
    line_spacing = line_spacing,
    border = border,
    padding.bottom = padding.bottom,
    padding.top = padding.top,
    padding.left = padding.left,
    padding.right = padding.right,
    border.bottom = border.bottom,
    border.left = border.left,
    border.top = border.top,
    border.right = border.right,
    shading.color = shading.color,
    keep_with_next = keep_with_next,
    tabs = tabs,
    word_style = word_style
  )
}


#' @export
#' @importFrom grDevices col2rgb
format.fp_par <- function(x, type = "wml", ...) {
  stopifnot(length(type) == 1)
  stopifnot(type %in% c("wml", "pml", "html", "rtf"))

  if (type == "wml") {
    ppr_wml(x)
  } else if (type == "pml") {
    ppr_pml(x)
  } else if (type == "html") {
    ppr_css(x)
  } else if (type == "rtf") {
    ppr_rtf(x)
  } else {
    stop("unimplemented")
  }
}
#' @export
to_wml.fp_par <- function(x, add_ns = FALSE, ...) {
  format(x, type = "wml")
}

#' @param x,object `fp_par` object
#' @param ... further arguments - not used
#' @rdname fp_par
#' @export
print.fp_par <- function(x, ...) {

  out <- data.frame(
    text.align = as.character(x$text.align),
    padding.top = as.character(x$padding.top),
    padding.bottom = as.character(x$padding.bottom),
    padding.left = as.character(x$padding.left),
    padding.right = as.character(x$padding.right),
    shading.color = as.character(x$shading.color)
  )
  out <- as.data.frame(t(out))
  names(out) <- "values"
  print(out)
  if (!is.null(x$border.top) && !isFALSE(x$border.top) &&
      !is.null(x$border.bottom) && !isFALSE(x$border.bottom) &&
      !is.null(x$border.left) && !isFALSE(x$border.left) &&
      !is.null(x$border.right) && !isFALSE(x$border.right)) {
    cat("borders:\n")
    borders <- rbind(
      as.data.frame(unclass(x$border.top)),
      as.data.frame(unclass(x$border.bottom)),
      as.data.frame(unclass(x$border.left)),
      as.data.frame(unclass(x$border.right))
    )
    row.names(borders) <- c("top", "bottom", "left", "right")
    print(borders)
  } else {
    cat("no borders!\n")
  }


}


#' @rdname fp_par
#' @examples
#' obj <- fp_par(text.align = "center", padding = 1)
#' update(obj, padding.bottom = 5)
#' @export
update.fp_par <- function(
  object,
  text.align,
  padding,
  border,
  padding.bottom,
  padding.top,
  padding.left,
  padding.right,
  border.bottom,
  border.left,
  border.top,
  border.right,
  shading.color,
  keep_with_next,
  word_style,
  ...
) {
  if (!missing(text.align)) {
    object <- check_set_choice(
      obj = object,
      value = text.align,
      choices = c("left", "right", "center", "justify")
    )
  }
  if (!missing(word_style)) {
    object <- check_set_chr(obj = object, word_style)
  }

  # padding checking
  if (!missing(padding)) {
    object <- check_spread_integer(
      object,
      padding,
      c(
        "padding.bottom",
        "padding.top",
        "padding.left",
        "padding.right"
      )
    )
  }
  if (!missing(padding.bottom)) {
    object <- check_set_integer(obj = object, padding.bottom)
  }
  if (!missing(padding.left)) {
    object <- check_set_integer(obj = object, padding.left)
  }
  if (!missing(padding.top)) {
    object <- check_set_integer(obj = object, padding.top)
  }
  if (!missing(padding.right)) {
    object <- check_set_integer(obj = object, padding.right)
  }

  # border checking
  if (!missing(border)) {
    object <- check_spread_border(
      obj = object,
      border,
      dest = c(
        "border.bottom",
        "border.top",
        "border.left",
        "border.right"
      )
    )
  }
  if (!missing(border.top)) {
    object <- check_set_border(obj = object, border.top)
  }
  if (!missing(border.bottom)) {
    object <- check_set_border(obj = object, border.bottom)
  }
  if (!missing(border.left)) {
    object <- check_set_border(obj = object, border.left)
  }
  if (!missing(border.right)) {
    object <- check_set_border(obj = object, border.right)
  }

  if (!missing(shading.color)) {
    object <- check_set_color(object, shading.color)
  }
  if (!missing(keep_with_next)) {
    object <- check_set_bool(object, keep_with_next)
  }

  object
}

# fp_cell ----

vertical.align.styles <- c("top", "center", "bottom")
text.directions <- c("lrtb", "tbrl", "btlr")

#' @title Cell formatting properties
#'
#' @description Create a `fp_cell` object that describes cell formatting properties.
#'
#' @param border shortcut for all borders.
#' @param border.bottom,border.left,border.top,border.right [fp_border()] for borders.
#' @param vertical.align cell content vertical alignment - a single character value,
#' expected value is one of "center" or "top" or "bottom"
#' @param margin shortcut for all margins.
#' @param margin.bottom,margin.top,margin.left,margin.right cell margins - 0 or positive integer value.
#' @param background.color cell background color - a single character value specifying a
#' valid color (e.g. "#000000" or "black").
#' @param text.direction cell text rotation - a single character value, expected
#' value is one of "lrtb", "tbrl", "btlr".
#' @param rowspan specify how many rows the cell is spanned over
#' @param colspan specify how many columns the cell is spanned over
#' @export
#' @family functions for defining formatting properties
fp_cell <- function(
  border = fp_border(width = 0),
  border.bottom,
  border.left,
  border.top,
  border.right,
  vertical.align = "center",
  margin = 0,
  margin.bottom,
  margin.top,
  margin.left,
  margin.right,
  background.color = "transparent",
  text.direction = "lrtb",
  rowspan = 1,
  colspan = 1
) {
  out <- list()

  # border checking
  out <- check_spread_border(
    obj = out,
    border,
    dest = c(
      "border.bottom",
      "border.top",
      "border.left",
      "border.right"
    )
  )
  if (!missing(border.top)) {
    out <- check_set_border(obj = out, border.top)
  }
  if (!missing(border.bottom)) {
    out <- check_set_border(obj = out, border.bottom)
  }
  if (!missing(border.left)) {
    out <- check_set_border(obj = out, border.left)
  }
  if (!missing(border.right)) {
    out <- check_set_border(obj = out, border.right)
  }

  # background-color checking
  out <- check_set_color(out, background.color)

  out <- check_set_choice(
    obj = out,
    value = vertical.align,
    choices = vertical.align.styles
  )
  out <- check_set_choice(
    obj = out,
    value = text.direction,
    choices = text.directions
  )

  # margin checking
  out <- check_spread_integer(
    out,
    margin,
    c(
      "margin.bottom",
      "margin.top",
      "margin.left",
      "margin.right"
    )
  )

  if (!missing(margin.bottom)) {
    out <- check_set_integer(obj = out, margin.bottom)
  }
  if (!missing(margin.left)) {
    out <- check_set_integer(obj = out, margin.left)
  }
  if (!missing(margin.top)) {
    out <- check_set_integer(obj = out, margin.top)
  }
  if (!missing(margin.right)) {
    out <- check_set_integer(obj = out, margin.right)
  }

  out <- check_set_integer(obj = out, rowspan)
  out <- check_set_integer(obj = out, colspan)

  class(out) <- "fp_cell"
  out
}


#' @export
#' @rdname fp_cell
#' @param x,object `fp_cell` object
#' @param type output type - one of 'wml', 'pml', 'html', 'rtf'.
#' @param ... further arguments - not used
format.fp_cell <- function(x, type = "wml", ...) {
  stopifnot(length(type) == 1)
  stopifnot(type %in% c("wml", "pml", "html", "rtf"))

  if (type == "wml") {
    tcpr_wml(x)
  } else if (type == "pml") {
    tcpr_pml(x)
  } else if (type == "html") {
    tcpr_css(x)
  } else if (type == "rtf") {
    tcpr_rtf(x)
  } else {
    stop("unimplemented")
  }
}

#' @export
to_wml.fp_cell <- function(x, add_ns = FALSE, ...) {
  format(x, type = "wml")
}

#' @export
#' @rdname fp_cell
print.fp_cell <- function(x, ...) {
  cat(format(x, type = "html"))
}


#' @rdname fp_cell
#' @examples
#' obj <- fp_cell(margin = 1)
#' update(obj, margin.bottom = 5)
#' @export
update.fp_cell <- function(
  object,
  border,
  border.bottom,
  border.left,
  border.top,
  border.right,
  vertical.align,
  margin = 0,
  margin.bottom,
  margin.top,
  margin.left,
  margin.right,
  background.color,
  text.direction,
  rowspan = 1,
  colspan = 1,
  ...
) {
  if (!missing(border)) {
    object <- check_spread_border(
      obj = object,
      border,
      dest = c(
        "border.bottom",
        "border.top",
        "border.left",
        "border.right"
      )
    )
  }
  if (!missing(border.top)) {
    object <- check_set_border(obj = object, border.top)
  }
  if (!missing(border.bottom)) {
    object <- check_set_border(obj = object, border.bottom)
  }
  if (!missing(border.left)) {
    object <- check_set_border(obj = object, border.left)
  }
  if (!missing(border.right)) {
    object <- check_set_border(obj = object, border.right)
  }

  # background-color checking
  if (!missing(background.color)) {
    object <- check_set_color(object, background.color)
  }

  if (!missing(vertical.align)) {
    object <- check_set_choice(
      obj = object,
      value = vertical.align,
      choices = vertical.align.styles
    )
  }
  if (!missing(text.direction)) {
    object <- check_set_choice(
      obj = object,
      value = text.direction,
      choices = text.directions
    )
  }

  # margin checking
  if (!missing(margin)) {
    object <- check_spread_integer(
      object,
      margin,
      c(
        "margin.bottom",
        "margin.top",
        "margin.left",
        "margin.right"
      )
    )
  }

  if (!missing(margin.bottom)) {
    object <- check_set_integer(obj = object, margin.bottom)
  }
  if (!missing(margin.left)) {
    object <- check_set_integer(obj = object, margin.left)
  }
  if (!missing(margin.top)) {
    object <- check_set_integer(obj = object, margin.top)
  }
  if (!missing(margin.right)) {
    object <- check_set_integer(obj = object, margin.right)
  }

  if (!missing(rowspan)) {
    object <- check_set_integer(obj = object, rowspan)
  }
  if (!missing(colspan)) {
    object <- check_set_integer(obj = object, colspan)
  }

  object
}
