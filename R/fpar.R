#' @export
#' @title concatenate formatted text
#' @description Create a paragraph representation by concatenating
#' formatted text or images. Modify default text and paragraph
#' formatting properties with update.
#' @details
#' \code{fortify_fpar}, \code{as.data.frame} are used internally and
#' are not supposed to be used by end user.
#'
#' @param fp_p paragraph formatting properties
#' @param fp_t default text formatting properties
#' @param x,object fpar object
#' @examples
#' fpar(ftext("hello", shortcuts$fp_bold()))
fpar <- function( ... ) {
  out <- list()
  out$chunks <- list(...)

  out$fp_p <- fp_par()
  out$fp_t <- fp_text()

  class(out) <- c("fpar")
  out
}

#' @export
#' @rdname fpar
#' @importFrom stats update
update.fpar <- function (object, fp_p=NULL, fp_t=NULL, ...){
  if(!is.null(fp_p)){
    object$fp_p <- fp_p
  }
  if(!is.null(fp_t)){
    object$fp_t <- fp_t
  }

  object
}


#' @export
#' @rdname fpar
fortify_fpar <- function(x){
  lapply(x$chunks, function(chk) {
    if( !inherits(chk, "cot") ){
      chk <- ftext(text = format(chk), prop = x$fp_t )
    }
    chk
  })
}


#' @rdname fpar
#' @export
as.data.frame.fpar <- function( x, ...){
  chks <- fortify_fpar(x)
  chks <- chks[sapply(chks, function(x) inherits(x, "ftext"))]
  chks <- mapply(function(x){
    data.frame(value = x$value, size = x$pr$font.size,
           bold = x$pr$bold, italic = x$pr$italic,
           font.family = x$pr$font.family, stringsAsFactors = FALSE )
  }, chks, SIMPLIFY = FALSE)
  rbind.match.columns(chks)
}

#' @export
#' @param type a string value ("pml", "wml" or "html").
#' @param ... unused
#' @rdname fpar
format.fpar <- function( x, type = "pml", ...){
  if( type != "console" )
    par_style <- format(x$fp_p, type = type)
  chks <- fortify_fpar(x)
  chks <- sapply(chks, format, type = type)
  chks <- paste0(chks, collapse = "")
  if( type == "pml" )
    out <- paste0("<a:p>", par_style, chks, "</a:p>")
  else if( type == "html" )
    out <- paste0("<p style=\"", par_style, "\">", chks, "</p>")
  else if( type == "wml" )
    out <- paste0("<w:p>", par_style, chks, "</w:p>")
  else if( type == "console" )
    out <- chks
  else stop("unimplemented")
  out
}

#' @export
#' @rdname fpar
print.fpar = function (x, ...){
  cat( format(x, type = "console"), "\n", sep = "" )
}

