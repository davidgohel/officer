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
  ptr <- rpr_pointer( x$pr )

  if( type == "wml" ){
    out <- chunk_w(x$value, ptr)
  } else if( type == "pml" ){
    out <- chunk_p(x$value, ptr)
  } else if( type == "html" ){
    out <- chunk_html(x$value, ptr)
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


#' @rdname minibar
#' @export
as.data.frame.external_img <- function( x, ... ){
  dimx <- attr(x, "dims")
  data.frame(path = as.character(x), width = dimx$width, height = dimx$height)
}

#' @importFrom openssl base64_encode
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
    input <- normalizePath(as.character(x), mustWork = TRUE)
    buf <- readBin(input, raw(), file.info(input)$size)
    base64 <- base64_encode(buf, linebreaks = FALSE)
    out <- sprintf("<img src=\"data:image/png;base64,\n%s\" width=\"%.0f\" height=\"%.0f\"/>",
            base64, dims$width*72, dims$height*72)
  } else if( type == "console" ){
    out <- paste0( "{image:{", as.character(x), "}}" )
  } else stop("unimplemented")
  out
}

print.external_img = function (x, ...){
  cat( format(x, type = "console"), "\n", sep = "" )
}



#' @export
#' @title draw a single bar
#' @description This function is used to insert bars into flextable with function \code{display}
#' @param value bar height
#' @param max max bar height
#' @param barcol bar color
#' @param bg background color
#' @param width,height size of the resulting png file in inches
#' @importFrom grDevices as.raster col2rgb rgb
minibar <- function(value, max, barcol = "#CCCCCC", bg = "transparent", width = 1, height = .2) {
  stopifnot(value >= 0, max >= 0)
  barcol <- rgb(t(col2rgb(barcol))/255)
  bg <- ifelse( bg == "transparent", bg, rgb(t(col2rgb(bg))/255) )
  if( value > max ){
    warning("value > max, truncate to max")
    value <- max
  }
  width_ <- as.integer(width * 72)
  value <- as.integer( (value / max) * width_ )
  n_empty <- width_ - value
  out <- tempfile(fileext = ".png")
  class(out) <- c("minibar", "cot")
  attr(out, "dims") <- list(width = width, height = height)
  attr(out, "r") <- as.raster( matrix(c(rep(barcol, value), rep(bg, n_empty)), nrow = 1) )

  out
}

#' @rdname minibar
#' @param x \code{minibar} object
#' @export
dim.minibar <- function( x ){
  x <- attr(x, "dims")
  data.frame(width = x$width, height = x$height)
}

#' @rdname minibar
#' @export
as.data.frame.minibar <- function( x, ... ){
  dimx <- attr(x, "dims")
  data.frame(path = as.character(x), width = dimx$width, height = dimx$height)
}

#' @param type output format
#' @param ... unused
#' @export
#' @importFrom gdtools raster_write raster_str
#' @rdname minibar
format.minibar = function (x, type = "console", ...){
  stopifnot( length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html", "console") )
  dims <- dim(x)
  width <- attr(x, "dims")$width * 72
  height <- attr(x, "dims")$height * 72
  r_ <- attr(x, "r")
  if( type == "pml" ){
    out <- ""
  } else if( type == "wml" ){
    file <- raster_write(x = r_, path = as.character(x), width = width, height = height )
    out <- wml_run_pic(file, width = width, height = height)
  } else if( type == "html" ){
    code <- raster_str(x = r_, width = width, height = height )
    out <- sprintf("<img src=\"data:image/png;base64,\n%s\" width=\"%.0f\" height=\"%.0f\"/>",
                   code, width, height)
  } else if( type == "console" ){
    out <- "{raster:{...}}"
  } else stop("unimplemented")
  out
}

