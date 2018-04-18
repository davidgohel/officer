#' @export
#' @title list of blocks
#' @description a list of blocks can be used to gather
#' several blocks (paragraphs or tables) into a single
#' object. The function is to be used when adding
#' footnotes for example.
#' @param ... a list of objects of class \code{fpar} or
#' \code{flextable}.
#' @examples
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' bl <- block_list(
#'   fpar(ftext("hello world", shortcuts$fp_bold())),
#'   fpar(
#'     ftext("hello", shortcuts$fp_bold()),
#'     stext(" world", "strong"),
#'     external_img(src = img.file, height = 1.06, width = 1.39)
#'   )
#' )
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


