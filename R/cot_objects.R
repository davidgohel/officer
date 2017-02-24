#' @export
#' @title formatted text
#' @description Format a chunk of text with text formatting properties.
#' @param text text value
#' @param prop formatting text properties
#' @examples
#' ftext("hello", fp_text())
ftext <- function(text, prop) {
  out <- list( value = formatC(text), pr = prop )
  class(out) <- c("ftext", "cot")
  out
}


#' @export
#' @param type output format, one of wml, pml, html, console, text.
#' @param ... unused
#' @rdname ftext
format.ftext = function (x, type = "console", ...){
  stopifnot( length(type) == 1,
             type %in% c("wml", "pml", "html", "console", "text") )

  if( type == "wml" ){
    out <- paste0("<w:r>", format(x$pr, type = type ),
           "<w:t xml:space=\"preserve\">",
           x$value, "</w:t></w:r>")
  } else if( type == "pml" ){
    out <- paste0("<a:r>", format(x$pr, type = type ),
                  "<a:t>", x$value, "</a:t></a:r>")
  } else if( type == "html" ){
    out <- paste0("<span style=\"", format(x$pr, type = type ),
                  "\">", x$value, "</span>")
  } else if( type == "console" ){
    out <- paste0( "{text:{", x$value, "}}" )
  } else if( type == "text" ){
    out <- x$value
  } else stop("unimplemented")

  out
}

#' @export
#' @rdname ftext
#' @param x \code{ftext} object
print.ftext = function (x, ...){
  cat( format(x, type = "console"), "\n", sep = "" )
}


#' @export
#' @title external image
#' @description This function is used to insert images into flextable with function \code{display}
#' @param src image file path
#' @param width height in inches
#' @param height height in inches
#' @examples
#' # external_img("example.png")
external_img <- function(src, width = .5, height = .2) {
  stopifnot( file.exists(src) )
  class(src) <- c("external_img", "cot")
  attr(src, "dims") <- list(width = width, height = height)
  src
}

#' @rdname external_img
#' @param x \code{external_img} object
#' @export
dim.external_img <- function( x ){
  x <- attr(x, "dims")
  data.frame(width = x$width, height = x$height)
}


#' @rdname external_img
#' @export
as.data.frame.external_img <- function( x, ... ){
  dimx <- attr(x, "dims")
  data.frame(path = as.character(x), width = dimx$width, height = dimx$height)
}

#' @importFrom base64enc dataURI
#' @param type output format
#' @param ... unused
#' @export
#' @rdname external_img
format.external_img = function (x, type = "console", ...){
  # r <- as.raster(c(0.5, 1, 0.5))
  # plot(r)
  stopifnot( length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html", "console") )
  dims <- dim(x)
  if( type == "pml" ){
    out <- pml_run_pic(as.character(x), width = dims$width*72, height = dims$height*72)
  } else if( type == "wml" ){
    out <- wml_run_pic(as.character(x), width = dims$width*72, height = dims$height*72)
  } else if( type == "html" ){
    out <- sprintf("<img src=\"%s\" width=\"%.0f\" height=\"%.0f\"/>",
                   dataURI(file = as.character(x), mime="image/png"),
                   dims$width*72, dims$height*72)
  } else if( type == "console" ){
    out <- paste0( "{image:{", as.character(x), "}}" )
  } else stop("unimplemented")
  out
}

print.external_img = function (x, ...){
  cat( format(x, type = "console"), "\n", sep = "" )
}



