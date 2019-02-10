#' @importFrom grDevices col2rgb rgb
colcode <- function(x) {
  rgb(t(col2rgb(x, alpha = TRUE)), maxColorValue = 255)
}
colalpha <- function(x){
  x[x%in%'transparent'] <- "#FFFFFF00"
  rgbvals <- col2rgb(x, alpha = TRUE)
  as.integer(rgbvals[4,] / 255 * 100000)
}

colcode0 <- function(x) {
  x[x%in%'transparent'] <- "#FFFFFF00"
  substr(rgb(t(col2rgb(x, alpha = TRUE)), maxColorValue = 255), 2, 7)
}
is.color = function(x) {
  # http://stackoverflow.com/a/13290832/3315962
  out = sapply(x, function( x ) {
    tryCatch( is.matrix( col2rgb( x ) ), error = function( e ) F )
  })

  nout <- names(out)
  if( !is.null(nout) && any( is.na( nout ) ) )
    out[is.na( nout )] = FALSE

  out
}

