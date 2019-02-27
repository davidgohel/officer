#' @export
#' @title unordered list
#' @description unordered list of text for PowerPoint
#' presentations. Each text is associated with
#' a hierarchy level.
#' @param str_list list of strings to be included in the object
#' @param level_list list of levels for hierarchy structure
#' @param style text style, a \code{fp_text} object list or a
#' single \code{fp_text} objects. Use \code{fp_text(font.size = 0, ...)} to
#' inherit from default sizes of the presentation.
#' @examples
#' unordered_list(
#' level_list = c(1, 2, 2, 3, 3, 1),
#' str_list = c("Level1", "Level2", "Level2", "Level3", "Level3", "Level1"),
#' style = fp_text(color = "red", font.size = 0) )
#' @seealso \code{\link{ph_with}}
unordered_list <- function(str_list = character(0), level_list = integer(0), style = NULL){
  stopifnot(is.character(str_list))
  stopifnot(is.numeric(level_list))

  if (length(str_list) != length(level_list) & length(str_list) > 0) {
    stop("str_list and level_list have different lenghts.")
  }

  if( !is.null(style)){
    if( inherits(style, "fp_text") )
      style <- lapply(seq_len(length(str_list)), function(x) style )
  }
  x <- list(
    str = str_list,
    lvl = level_list,
    style = style
  )
  class(x) <- "unordered_list"
  x
}
#' @export
#' @noRd
print.unordered_list <- function(x, ...){
  print(data.frame(str = x$str,
                   lvl = x$lvl,
                   stringsAsFactors = FALSE))
  invisible()
}


