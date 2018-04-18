#' @export
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

#' @export
add_block <- function(x, value ){
  x
}


