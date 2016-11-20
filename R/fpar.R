#' @importFrom lazyeval f_list
#' @export
#' @title formatted paragraph
#' @description formatted paragraph
#' @param fp_p paragraph formatting properties
#' @param fp_t default text formatting properties
#' @param x,object fpar object
#' @examples
#' fpar(ftext("hello", fp_bold()))
fpar <- function( ... ) {
  out <- list()
  out$chunks <- f_list(...)

  out$fp_p <- fp_par()
  out$fp_t <- fp_text()

  class(out) <- c("fpar")
  out
}

#' @export
#' @rdname fpar
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
#' @importFrom purrr map_df
dim.fpar <- function( x ){
  padding.left <- x$fp_p$padding.left * (4/3) / 72
  padding.right <- x$fp_p$padding.right * (4/3) / 72
  padding.top <- x$fp_p$padding.top * (4/3) / 72
  padding.bottom <- x$fp_p$padding.bottom * (4/3) / 72

  out <- list( width = 0, height = 0 )
  if( length(x$chunks) > 0 ){
    chks <- cast_chunks(x)
    x <- map_df(chks, dim )
    out <- list( width = sum( x$width ),
               height = max( x$height ) )
  }

  out$width <- out$width + ( padding.left + padding.right)
  out$height <- out$height + ( padding.top + padding.bottom)
  out
}

#' @importFrom purrr map
#' @export
cast_chunks <- function(x){
  map(x$chunks, function(chk) {
    if( !inherits(chk, "cot") ){
      if( is.character(chk) )
        chk <- ftext(text = chk, prop = x$fp_t )
      else if( is.factor(chk) )
        chk <- ftext(text = as.character(chk), prop = x$fp_t )
      else chk <- ftext(text = format(chk), prop = x$fp_t )
    }
    chk
  })
}

#' @export
#' @param type a string value ("pml", "wml" or "html").
#' @param ... unused
#' @importFrom purrr map_chr
#' @rdname fpar
format.fpar <- function( x, type = "pml", ...){
  if( type != "console" )
    par_style <- format(x$fp_p, type = type)
  chks <- cast_chunks(x)
  if( inherits(try(expr = map_chr(chks, format, type = type) ), "try-error"))
    browser()
  chks <- map_chr(chks, format, type = type)
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

set_rid <- function( x, id ){
  for(i in seq_along(x$chunks) ){
    if( inherits(x$chunks[[i]], "external_img")){
      attr(x$chunks[[i]], "id") <- sprintf("rId%0.f", id + i )
    }
  }
  x
}

#' @export
get_imgs <- function( x ){
  out <- character(0)
  for(i in seq_along(x$chunks) ){
    if( inherits(x$chunks[[i]], "external_img")){
      out <- append(out, as.character(x$chunks[[i]]))
    } else if( inherits(x$chunks[[i]], "minibar")){
      out <- append(out, as.character(x$chunks[[i]]))
    }
  }
  out
}
