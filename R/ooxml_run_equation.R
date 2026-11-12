# equations ----

#' @export
#' @title Equation
#' @description Create an equation written in LaTeX, to be used in a
#' paragraph made with [fpar()].
#'
#' The equation is converted by the package 'equatags' into an equation of
#' 'Word' ('Office Open XML Math'): it stays editable in 'Word'.
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
#' the default size and color of the document.
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
