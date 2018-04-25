#' @export
#' @title concatenate formatted text
#' @description Create a paragraph representation by concatenating
#' formatted text or images.
#'
#' \code{fpar} supports \code{ftext}, \code{external_img} and simple strings.
#' All its arguments will be concatenated to create a paragraph where chunks
#' of text and images are associated with formatting properties.
#'
#' Default text and paragraph formatting properties can also
#' be modified with update.
#' @details
#' \code{fortify_fpar}, \code{as.data.frame} are used internally and
#' are not supposed to be used by end user.
#'
#' @param fp_p paragraph formatting properties
#' @param fp_t default text formatting properties. This is used as
#' text formatting properties when simple text is provided as argument.
#'
#' @param x,object fpar object
#' @examples
#' fpar(ftext("hello", shortcuts$fp_bold()))
#'
#' # mix text and image -----
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#'
#' bold_face <- shortcuts$fp_bold(font.size = 12)
#' bold_redface <- update(bold_face, color = "red")
#' fpar_1 <- fpar(
#'   "Hello World, ",
#'   ftext("how ", prop = bold_redface ),
#'   external_img(src = img.file, height = 1.06/2, width = 1.39/2),
#'   ftext(" you?", prop = bold_face ) )
#' fpar_1
#'
#' img_in_par <- fpar(
#'   external_img(src = img.file, height = 1.06/2, width = 1.39/2),
#'   fp_p = fp_par(text.align = "center") )
fpar <- function( ..., fp_p = fp_par(), fp_t = fp_text() ) {
  out <- list()
  out$chunks <- list(...)
  out$fp_p <- fp_p
  out$fp_t <- fp_t
  class(out) <- c("fpar")
  out
}

#' @export
#' @rdname fpar
#' @importFrom stats update
update.fpar <- function (object, fp_p = NULL, fp_t = NULL, ...){

  if(!is.null(fp_p)){
    object$fp_p <- fp_p
  }
  if(!is.null(fp_t)){
    object$fp_t <- fp_t
  }

  object
}


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
print.fpar = function (x, ...){
  cat( format(x, type = "console"), "\n", sep = "" )
}

