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
           htmlEscape(x$value), "</w:t></w:r>")
  } else if( type == "pml" ){
    out <- paste0("<a:r>", format(x$pr, type = type ),
                  "<a:t>", htmlEscape(x$value), "</a:t></a:r>")
  } else if( type == "html" ){
    out <- paste0("<span style=\"", format(x$pr, type = type ),
                  "\">", htmlEscape(x$value), "</span>")
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
#' @description This function is used to insert images in 'PowerPoint'
#' slides.
#' @param src image file path
#' @param width height in inches
#' @param height height in inches
#' @examples
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#'
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc,
#'   value = external_img(img.file, width = 1.39, height = 1.06),
#'   location = ph_location_type(type = "body"),
#'   use_loc_size = FALSE )
#' print(doc, target = tempfile(fileext = ".pptx"))
#' @seealso [ph_with]
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
  stopifnot( length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html", "console") )
  dims <- dim(x)
  if( type == "pml" ){
    out <- pml_image(as.character(x), width = dims$width, height = dims$height)
  } else if( type == "wml" ){
    out <- wml_image(src=as.character(x), width = dims$width, height = dims$height)
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

