#' @title Wrap plot instructions for png plotting in Powerpoint or Word
#' @description A simple wrapper to capture
#' plot instructions that will be executed and copied in a document. It produces
#' an object of class 'plot_instr' with a corresponding method [ph_with()].
#'
#' The function enable usage of any R plot with argument `code`. Wrap your code
#' between curly bracket if more than a single expression.
#'
#' @param code plotting instructions
#' @examples
#' anyplot <- plot_instr(code = barplot(1:5, col = 2:6))
#'
#' library(officer)
#' doc <- read_pptx()
#' doc <- add_slide(doc, "Title and Content", "Office Theme")
#' doc <- ph_with(doc, anyplot, location = ph_location_fullsize())
#' print(doc, target = tempfile(fileext = ".pptx"))
#' @export
#' @import graphics
#' @seealso \code{\link{ph_with}}
plot_instr <- function(code) {
  out <- list()
  out$code <- substitute(code)
  class(out) <- "plot_instr"
  return(out)
}
