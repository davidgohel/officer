#' @export
#' @title create paragraph blocks
#' @description a list of blocks can be used to gather
#' several blocks (paragraphs or tables) into a single
#' object. The function is to be used when adding
#' footnotes or formatted paragraphs into a new slide.
#' @param ... a list of objects of class \code{\link{fpar}} or
#' \code{flextable}. When output is only for Word, objects
#' of class \code{\link{external_img}} can also be used in
#' fpar construction to mix text and images in a single paragraph.
#' @examples
#'
#' @example examples/block_list.R
#' @seealso [ph_with()], [body_add_blocks()]
block_list <- function(...){
  x <- list(...)
  class(x) <- "block_list"
  x
}
#' @export
print.block_list <- function(x, ...){
  for( i in x ){
    cat(format(i, type = "console"), "\n")
  }
  invisible()
}


