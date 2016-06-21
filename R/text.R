#' @export
ftext <- function(text, prop) {
  out <- list( value = text,
               pr = prop )
  class(out) <- c("ftext", "cot")
  out
}

#' @export
format.ftext = function (x, type = "wml", ...){

  stopifnot( length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html") )
  if( type == "wml" ){
    text_style <- format(x$pr, type = type)
    out <- paste0("<w:r>", text_style, "<w:t>", x$value, "</w:t></w:r>")

  } else if( type == "pml" ){
    text_style <- format(x$pr, type = type)
    out <- paste0("<a:r>", text_style, "<a:t>", x$value, "</a:t></a:r>")

  } else if( type == "html" ){

  } else stop("unimplemented")
  out
}

#' @export
fraster <- function(x, height = .2, width = .5) {
  stopifnot( inherits(x, "raster") )
  out <- list( x )
  class(out) <- c("fraster", "cot")
  out
}
#' @export
format.fraster = function (x, type = "wml", ...){
  # r <- as.raster(c(0.5, 1, 0.5))
  # plot(r)
  stopifnot( length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html") )
  if( type == "wml" ){

  } else if( type == "pml" ){

  } else if( type == "html" ){

  } else stop("unimplemented")
  out
}



#' @export
#' @title paragraph container
#' @description a paragraph container
#' @importFrom R6 R6Class
#' @importFrom purrr map_chr
paragraph <- R6Class(
  "paragraph",
  public = list(
    initialize = function(prop) {
      private$chunks <- list()
      private$pr <- prop
    },
    add = function(x){
      if( !inherits(x, "cot") ){
        stop("x should be an object of class cot")
      }
      private$chunks <- append(private$chunks, list(x))
      self
    },
    format = function(type = "pml"){
      par_style <- format(private$pr, type = type)
      chks <- map_chr(private$chunks, format, type = type)
      chks <- paste0(chks, collapse = "")
      out <- paste0("<a:p>", par_style, chks, "</a:p>")
      out
    }
  ),
  private = list(
    chunks = NULL,
    pr = NULL
  )
)


