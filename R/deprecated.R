#' @export
#' @title append a footnote
#' @description this function is deprecated.
#' @param x an rdocx object
#' @param ... unused
#' @note This function is now deprecated because it is not
#' efficient and make users write complex code. Use instead [fpar()] to build
#' formatted paragraphs.
#' @keywords internal
slip_in_footnote <- function(x, ...){
  message("`slip_in_footnote()` is deprecated, it is replaced by `run_footnote()`. Pausing for 10 seconds.")
  Sys.sleep(10)
  x
}
