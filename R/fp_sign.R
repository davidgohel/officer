#' @import digest digest
#'
#' @title object unique signature
#' @description Get unique signature for a formatting properties
#' object.
#' @param x a set of formatting properties
#' @examples
#' fp_sign( fp_text(color="orange") )
#' @export
fp_sign <- function( x ){
  digest(x)
}


