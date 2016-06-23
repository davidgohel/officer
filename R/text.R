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
    text_style <- format(x$pr, type = type)
    out <- paste0("<span style=\"", text_style, "\">", x$value, "</span>")

  } else stop("unimplemented")
  out
}

#' @export
#' @importFrom gdtools str_extents
dim.ftext <- function( x ){
  mat <- str_extents(x = x$value,
                     fontname = x$pr$font.family,
                     fontsize = x$pr$font.size,
                     bold = x$pr$bold,
                     italic = x$pr$italic) / 72
  mat <- as.data.frame(mat)
  names(mat ) <- c("width", "height")
  mat
}


#' @export
external_img <- function(x, id=1, height = .2, width = .5) {
  stopifnot( file.exists(x) )
  class(x) <- c("external_img", "cot")
  attr(x, "rid") <- id
  attr(x, "dims") <- list(width = width, height = height)
  x
}
#' @export
dim.external_img <- function( x ){
  x <- attr(x, "dims")
  data.frame(width = x$width, height = x$height)
}


#' @importFrom openssl base64_encode
#' @export
format.external_img = function (x, type = "wml", ...){
  # r <- as.raster(c(0.5, 1, 0.5))
  # plot(r)
  stopifnot( length(type) == 1)
  stopifnot( type %in% c("wml", "pml", "html") )
  if( type == "pml" ){
    out <- ""
  } else if( type == "html" ){
    dims <- dim(x)
    input <- normalizePath(as.character(x), mustWork = TRUE)
    buf <- readBin(input, raw(), file.info(input)$size)
    base64 <- base64_encode(buf, linebreaks = FALSE)
    out <- sprintf("<img src=\"data:image/png;base64,\n%s\" width=\"%.0f\" height=\"%.0f\"/>",
            base64, dims$width*72, dims$height*72)
  } else stop("unimplemented")
  out
}



#' @export
#' @title paragraph container
#' @description a paragraph container
#' @importFrom R6 R6Class
#' @importFrom purrr map_chr map_df
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
    set_rid = function( id ){
      for(i in seq_along(private$chunks) ){
        if( inherits(private$chunks[[i]], "external_img")){
          attr(private$chunks[[i]], "id") <- sprintf("rId%0.f", id + i )
        }
      }
      self
    },
    get_imgs = function(  ){
      out <- character(0)
      for(i in seq_along(private$chunks) ){
        if( inherits(private$chunks[[i]], "external_img")){
          out <- append(out, as.character(private$chunks[[i]]))
        }
      }
      out
    },
    dim = function( ){
      as.list( colSums( map_df(private$chunks, dim ) ) )
    },
    format = function(type = "pml"){
      par_style <- format(private$pr, type = type)
      chks <- map_chr(private$chunks, format, type = type)
      chks <- paste0(chks, collapse = "")
      if( type == "pml" )
        out <- paste0("<a:p>", par_style, chks, "</a:p>")
      else if( type == "html" )
        out <- paste0("<p style=\"", par_style, "\">", chks, "</p>")
      else stop("unimplemented")
      out
    }
  ),
  private = list(
    chunks = NULL,
    pr = NULL
  )
)

