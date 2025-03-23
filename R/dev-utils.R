#' @importFrom openssl base64_encode
#' @export
#' @title Images to base64
#' @description encodes images into base64 strings.
#' @param filepaths file names.
#' @keywords internal
#' @examples
#' rlogo <- file.path( R.home("doc"), "html", "logo.jpg")
#' image_to_base64(rlogo)
image_to_base64 <- function(filepaths){
  vapply(
    filepaths,
    function(filepath){
      if( grepl("\\.png", ignore.case = TRUE, x = filepath) ){
        mime <- "image/png"
      } else if( grepl("\\.gif", ignore.case = TRUE, x = filepath) ){
        mime <- "image/gif"
      } else if( grepl("\\.jpg", ignore.case = TRUE, x = filepath) ){
        mime <- "image/jpeg"
      } else if( grepl("\\.jpeg", ignore.case = TRUE, x = filepath) ){
        mime <- "image/jpeg"
      } else if( grepl("\\.svg", ignore.case = TRUE, x = filepath) ){
        mime <- "image/svg+xml"
      } else if( grepl("\\.tiff", ignore.case = TRUE, x = filepath) ){
        mime <- "image/tiff"
      } else if( grepl("\\.pdf", ignore.case = TRUE, x = filepath) ){
        mime <- "application/pdf"
      } else if( grepl("\\.webp", ignore.case = TRUE, x = filepath) ){
        mime <- "image/webp"
      } else {
        stop(sprintf("'officer' does not know how to encode format of the file '%s'.", filepath))
      }
      if(!file.exists(filepath)){
        stop(sprintf("file '%s' can not be found.",filepath))
      }
      dat <- readBin(filepath, what = "raw", size = 1, endian = "little", n = 1e+8)
      base64_str <- base64_encode(bin = dat)
      base64_str <- sprintf("data:%s;base64,%s", mime, base64_str)
      base64_str
    },
    FUN.VALUE = "")
}

#' @importFrom uuid UUIDgenerate
#' @export
#' @title generates unique identifiers
#' @description generates unique identifiers based
#' on [uuid::UUIDgenerate()].
#' @param n integer, number of unique identifiers to generate.
#' @param ... arguments sent to [uuid::UUIDgenerate()]
#' @keywords internal
#' @examples
#' uuid_generate(n = 5)
uuid_generate <- function(n = 1, ...) {
  UUIDgenerate(n = n, ...)
}

.url_special_chars <- list(
  `&` = '&amp;',
  `<` = '&lt;',
  `>` = '&gt;',
  `'` = '&#39;',
  `"` = '&quot;',
  ` ` = "&nbsp;"
)

#' @export
#' @title officer url encoder
#' @description encode url so that it can be easily
#' decoded when 'officer' write a file to the disk.
#' @param x a character vector of URL
#' @keywords internal
#' @examples
#' officer_url_encode("https://cran.r-project.org/")
officer_url_encode <- function(x) {
  for (chr in names(.url_special_chars)) {
    x <- gsub(chr, .url_special_chars[[chr]], x, fixed = TRUE, useBytes = TRUE)
  }
  Encoding(x) <- "UTF-8"
  x
}

officer_url_decode <- function(x) {
  for (chr in rev(names(.url_special_chars))) {
    x <- gsub(.url_special_chars[[chr]], chr, x, fixed = TRUE, useBytes = TRUE)
  }
  Encoding(x) <- "UTF-8"
  x
}
