# equations ----

#' @export
#' @title Equation
#' @description Create an equation written in LaTeX, to be used in a
#' paragraph made with [fpar()].
#'
#' The equation is converted by the package 'equatags' into an 'Office'
#' equation ('Office Open XML Math'): it stays editable in 'Word' and
#' 'PowerPoint'. Use it with [body_add_fpar()] in a 'Word' document, with
#' [ph_with()] in a 'PowerPoint' presentation and with [sheet_add_drawing()]
#' in an 'Excel' workbook, where equations are in a text box ('Excel' does
#' not accept equations in cells).
#'
#' In HTML (`to_html()`), equations are drawn by 'KaTeX' (its stylesheet is
#' linked). In RTF documents ([rtf_add()]), they are images drawn by
#' 'equatags' (R >= 4.3), placed on the baseline of the text; with an older
#' version of R, the LaTeX source is written as text, with a warning.
#'
#' @section Inline and display style:
#' As in 'Word', the style of the equation depends on its paragraph:
#'
#' - an equation alone in its paragraph uses the display style (larger
#'   fractions, limits above and below sums) and is aligned with the
#'   alignment of the paragraph (argument `text.align` of [fp_par()]);
#' - an equation in a paragraph that contains something else, even a space,
#'   uses the inline style.
#'
#' @inheritSection ftext usage
#' @param x the equation, a LaTeX string. The functions supported by 'KaTeX'
#' can be used, see <https://katex.org/docs/supported.html>.
#' @param prop formatting text properties returned by [fp_text()] or
#' [fp_text_lite()]. Only the font size and the color are used, the font is
#' the math font of the document ('Cambria Math' by default). With `NULL`,
#' the default size and color of the document or of the shape.
#' @examples
#' library(officer)
#'
#' quadratic <- "x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}"
#' euler <- "e^{i\\pi} + 1 = 0"
#'
#' doc <- read_docx()
#' # an equation alone in its paragraph: display style, centered
#' doc <- body_add_fpar(
#'   doc,
#'   fpar(run_equation(quadratic), fp_p = fp_par(text.align = "center"))
#' )
#' # an equation in a sentence: inline style, 14 points, in red
#' doc <- body_add_fpar(
#'   doc,
#'   fpar(
#'     "Euler's identity: ",
#'     run_equation(euler, prop = fp_text_lite(font.size = 14, color = "red")),
#'     "."
#'   )
#' )
#' print(doc, target = tempfile(fileext = ".docx"))
#'
#' pres <- read_pptx()
#' pres <- add_slide(pres, layout = "Title and Content")
#' pres <- ph_with(
#'   pres,
#'   block_list(
#'     fpar("The quadratic formula:"),
#'     fpar(run_equation(quadratic), fp_p = fp_par(text.align = "center"))
#'   ),
#'   location = ph_location_type("body")
#' )
#' print(pres, target = tempfile(fileext = ".pptx"))
#'
#' wb <- read_xlsx()
#' wb <- sheet_add_drawing(
#'   wb,
#'   value = fpar(run_equation(quadratic)),
#'   sheet = "Feuil1",
#'   anchor = "C2"
#' )
#' print(wb, target = tempfile(fileext = ".xlsx"))
#' @family run functions for reporting
run_equation <- function(x, prop = NULL) {
  check_equatags()
  if (!is.character(x) || length(x) != 1L || is.na(x)) {
    stop("`x` must be a single LaTeX string.", call. = FALSE)
  }
  if (!is.null(prop) && !inherits(prop, "fp_text")) {
    stop("`prop` must be NULL or an object made by fp_text().", call. = FALSE)
  }
  z <- list(value = x, pr = prop)
  class(z) <- c("run_equation", "run")
  z
}

#' @export
print.run_equation <- function(x, ...) {
  cat("equation:", x$value, "\n")
  invisible(x)
}

#' @export
format.run_equation <- function(x, ...) {
  x$value
}

check_equatags <- function() {
  if (
    !requireNamespace("equatags", quietly = TRUE) ||
      utils::packageVersion("equatags") < "0.4.0"
  ) {
    stop(
      "Package 'equatags' (>= 0.4.0) is required for equations: ",
      "install.packages(\"equatags\").",
      call. = FALSE
    )
  }
}

# Run properties of the math runs: only the size and the color, the font of
# an equation is the math font of the document.
equation_rpr_wml <- function(pr) {
  if (is.null(pr)) {
    return(NULL)
  }
  out <- paste0(
    if (!is.na(pr$font.size)) {
      sprintf(
        "<w:sz w:val=\"%.0f\"/><w:szCs w:val=\"%.0f\"/>",
        pr$font.size * 2,
        pr$font.size * 2
      )
    },
    if (!is.na(pr$color)) {
      sprintf("<w:color w:val=\"%s\"/>", hex_color(pr$color))
    }
  )
  if (!nzchar(out)) {
    return(NULL)
  }
  paste0("<w:rPr>", out, "</w:rPr>")
}

equation_omml <- function(x, display, rpr) {
  out <- if (is.null(rpr)) {
    equatags::transform_mathjax(x, to = "mml", display = display)
  } else {
    equatags::transform_mathjax(x, to = "mml", display = display, rpr = rpr)
  }
  trimws(out)
}

#' @export
to_wml.run_equation <- function(x, add_ns = FALSE, display = FALSE, ...) {
  equation_omml(x$value, display = display, rpr = equation_rpr_wml(x$pr))
}

is_equation_par <- function(chunks) {
  length(chunks) > 0L &&
    all(vapply(chunks, inherits, NA, what = "run_equation"))
}

# Justification of a math paragraph from the alignment of its paragraph.
math_jc <- function(text.align) {
  jc <- c(left = "left", right = "right", center = "center", justify = "left")
  out <- unname(jc[text.align])
  if (length(out) != 1L || is.na(out)) "centerGroup" else out
}

# A paragraph that only contains equations, as Word writes it: a math
# paragraph (m:oMathPara), equations in display style.
equation_par_wml <- function(chunks, fp_p) {
  eqs <- vapply(
    chunks,
    function(chk) to_wml(chk, display = TRUE),
    FUN.VALUE = ""
  )
  paste0(
    "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">",
    "<m:oMathParaPr><m:jc m:val=\"",
    math_jc(fp_p$text.align),
    "\"/></m:oMathParaPr>",
    paste0(eqs, collapse = ""),
    "</m:oMathPara>"
  )
}

# PowerPoint ----

equation_rpr_pml <- function(pr) {
  if (is.null(pr)) {
    return(NULL)
  }
  size <- if (!is.na(pr$font.size)) {
    sprintf(" sz=\"%.0f\"", pr$font.size * 100)
  } else {
    ""
  }
  fill <- if (!is.na(pr$color)) {
    alpha <- colalpha(pr$color)
    sprintf(
      "<a:solidFill><a:srgbClr val=\"%s\">%s</a:srgbClr></a:solidFill>",
      hex_color(pr$color),
      if (alpha < 1) sprintf("<a:alpha val=\"%.0f\"/>", alpha * 100000) else ""
    )
  } else {
    ""
  }
  if (!nzchar(size) && !nzchar(fill)) {
    return(NULL)
  }
  paste0("<a:rPr", size, ">", fill, "</a:rPr>")
}

# The equations of PowerPoint (and of the text boxes of Excel) are in an
# a14:m element, read by Office 2010 and later, inside mc:AlternateContent.
# As PowerPoint writes them: an inline equation is an m:oMath, the equations
# of a paragraph that only contains equations are in an m:oMathPara.
a14_math_pml <- function(x) {
  paste0(
    "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">",
    "<mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\">",
    "<a14:m>",
    x,
    "</a14:m></mc:Choice></mc:AlternateContent>"
  )
}

#' @export
to_pml.run_equation <- function(x, add_ns = FALSE, display = FALSE, ...) {
  omml <- equation_omml(
    x$value,
    display = display,
    rpr = equation_rpr_pml(x$pr)
  )
  if (display) {
    return(omml)
  }
  a14_math_pml(omml)
}

equation_par_pml <- function(chunks, fp_p) {
  eqs <- vapply(
    chunks,
    function(chk) to_pml(chk, display = TRUE),
    FUN.VALUE = ""
  )
  a14_math_pml(paste0(
    "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">",
    "<m:oMathParaPr><m:jc m:val=\"",
    math_jc(fp_p$text.align),
    "\"/></m:oMathParaPr>",
    paste0(eqs, collapse = ""),
    "</m:oMathPara>"
  ))
}

# Excel ----

# Size (inches) of a text box that only contains paragraphs of equations,
# from the size computed by equatags::equation_size() (font 'STIX Two Math',
# a little larger than 'Cambria Math'), the paddings of the paragraphs and
# the insets of the text box (0.1 inch left and right, 0.05 inch top and
# bottom); NULL when a paragraph contains something else. Without `prop`,
# the size of the equations is the default size of a text box in Excel, 11
# points.
equation_textbox_size <- function(value) {
  pars <- lapply(value, fortify_fpar)
  if (!length(pars) || !all(vapply(pars, is_equation_par, NA))) {
    return(NULL)
  }
  sizes <- mapply(
    function(chunks, fp_p) {
      s <- do.call(rbind, lapply(chunks, function(chk) {
        font_size <- if (is.null(chk$pr) || is.na(chk$pr$font.size)) {
          11
        } else {
          chk$pr$font.size
        }
        equatags::equation_size(chk$value, size = font_size, display = TRUE)
      }))
      pad <- function(p) if (is.na(p)) 0 else p / 72
      c(
        width = sum(s$width) + pad(fp_p$padding.left) + pad(fp_p$padding.right),
        height = max(s$height + s$depth) * 1.2 +
          pad(fp_p$padding.top) + pad(fp_p$padding.bottom)
      )
    },
    pars,
    lapply(value, function(z) z$fp_p)
  )
  list(
    width = max(sizes["width", ]) + 0.2,
    height = sum(sizes["height", ]) + 0.1
  )
}

# HTML ----

# KaTeX draws equations at 1.21 em (`.katex {font-size: 1.21em}`): the font
# size of `prop` is divided by 1.21 so that equations have this size. An
# equation alone in its paragraph follows the alignment and the spacing of
# the paragraph (KaTeX centers them, with margins).
#' @export
to_html.run_equation <- function(x, display = FALSE, ...) {
  out <- equatags::transform_mathjax(x$value, to = "html", display = display)
  out <- trimws(out)
  if (display) {
    out <- sub(
      "<span class=\"katex-display\"><span class=\"katex\">",
      paste0(
        "<span class=\"katex-display\" style=\"text-align:inherit;margin:0;\">",
        "<span class=\"katex\" style=\"text-align:inherit;\">"
      ),
      out,
      fixed = TRUE
    )
  }
  css <- character()
  if (!is.null(x$pr) && !is.na(x$pr$font.size)) {
    css <- c(css, sprintf("font-size:%.2fpt;", x$pr$font.size / 1.21))
  }
  if (!is.null(x$pr) && !is.na(x$pr$color)) {
    css <- c(css, sprintf("color:#%s;", hex_color(x$pr$color)))
  }
  if (length(css)) {
    out <- sprintf("<span style=\"%s\">%s</span>", paste0(css, collapse = ""), out)
  }
  out
}

# RTF ----

can_draw_equations <- function() {
  getRversion() >= "4.3.0" &&
    "equation_grob" %in% getNamespaceExports("equatags")
}

# RTF has no equations readable by every reader: an equation is a PNG image
# drawn by equatags::equation_grob() (R >= 4.3), lowered by its depth
# (`\dn`, in half points) so that its baseline is the baseline of the text.
# Without the size of `prop`, 11 points.
#' @export
to_rtf.run_equation <- function(x, display = FALSE, ...) {
  if (!can_draw_equations()) {
    warning(
      "Equations are written as LaTeX text in RTF: drawing them needs ",
      "R >= 4.3.",
      call. = FALSE
    )
    return(str_encode_to_rtf(x$value))
  }
  size <- if (is.null(x$pr) || is.na(x$pr$font.size)) 11 else x$pr$font.size
  colour <- if (is.null(x$pr) || is.na(x$pr$color)) "black" else x$pr$color
  s <- equatags::equation_size(x$value, size = size, display = display)
  height <- s$height + s$depth
  if (is.na(s$width) || s$width <= 0 || height <= 0) {
    return(str_encode_to_rtf(x$value))
  }
  path <- tempfile(fileext = ".png")
  on.exit(unlink(path))
  ragg::agg_png(
    filename = path,
    width = s$width,
    height = height,
    units = "in",
    res = 300,
    background = "transparent"
  )
  grid::grid.newpage()
  grid::grid.draw(equatags::equation_grob(
    x$value,
    size = size,
    display = display,
    x0 = 0,
    y0 = 0,
    hjust = 0,
    vjust = 0,
    colour = colour
  ))
  grDevices::dev.off()
  img <- to_rtf(external_img(src = path, width = s$width, height = height))
  lower <- round(s$depth * 72 * 2)
  if (lower > 0) {
    img <- sprintf("{\\dn%.0f %s}", lower, img)
  }
  img
}
