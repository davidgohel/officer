# Dimensions in pixels read from the header of the file, or NULL if the format
# is not recognised or the file is malformed: the image is not decoded.
image_dims_from_header <- function(path) {
  con <- file(path, "rb")
  on.exit(close(con))
  header <- readBin(con, "raw", n = 26L)
  if (length(header) < 10L) {
    return(NULL)
  }
  uint <- function(bytes) sum(as.numeric(bytes) * 256^(rev(seq_along(bytes)) - 1))

  png_signature <- as.raw(c(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a))
  if (identical(header[1:8], png_signature)) {
    if (length(header) < 24L || rawToChar(header[13:16]) != "IHDR") {
      return(NULL)
    }
    return(c(width = uint(header[17:20]), height = uint(header[21:24])))
  }
  if (identical(header[1:3], charToRaw("GIF"))) {
    return(c(width = uint(header[8:7]), height = uint(header[10:9])))
  }
  if (identical(header[1:3], as.raw(c(0xff, 0xd8, 0xff)))) {
    return(jpeg_dims_from_header(con, uint))
  }
  NULL
}

# walk the JPEG segments up to the frame header (SOF0 to SOF15, except DHT,
# JPG and DAC, which share the range)
jpeg_dims_from_header <- function(con, uint) {
  seek(con, 2L, origin = "start")
  repeat {
    prefix <- readBin(con, "raw", n = 1L)
    if (length(prefix) != 1L || prefix != as.raw(0xff)) {
      return(NULL)
    }
    marker <- readBin(con, "raw", n = 1L)
    while (length(marker) == 1L && marker == as.raw(0xff)) {
      marker <- readBin(con, "raw", n = 1L)
    }
    if (length(marker) != 1L) {
      return(NULL)
    }
    code <- as.integer(marker)
    # no payload for restart markers and TEM; start of scan: no frame found
    if (code == 0x01 || (code >= 0xd0 && code <= 0xd7)) {
      next
    }
    if (code == 0xd9 || code == 0xda) {
      return(NULL)
    }
    size <- readBin(con, "raw", n = 2L)
    if (length(size) != 2L || uint(size) < 2) {
      return(NULL)
    }
    if (code >= 0xc0 && code <= 0xcf && !code %in% c(0xc4, 0xc8, 0xcc)) {
      frame <- readBin(con, "raw", n = 5L)
      if (length(frame) != 5L) {
        return(NULL)
      }
      return(c(width = uint(frame[4:5]), height = uint(frame[2:3])))
    }
    seek(con, uint(size) - 2, origin = "current")
  }
}

#' @export
#' @title Get the dimensions of an image
#' @description Get the width and height in pixels of an image file or of a
#' raster object.
#'
#' The dimensions of 'png', 'jpeg' and 'gif' files are read from their header,
#' without decoding the image, which is much faster. Other formats are read with
#' package 'magick', which must then be installed.
#' @param x path to an image file, or a raster object (see [grDevices::as.raster()]).
#' @return A named numeric vector: `width` and `height`, in pixels.
#' @examples
#' img <- file.path(R.home("doc"), "html", "logo.jpg")
#' image_dims(img)
#'
#' image_dims(as.raster(matrix(0:1, nrow = 2, ncol = 3)))
#' @keywords internal
image_dims <- function(x) {
  if (grDevices::is.raster(x) || inherits(x, "nativeRaster")) {
    return(c(width = ncol(x), height = nrow(x)))
  }
  stopifnot(is.character(x), length(x) == 1L)

  dims <- NULL
  if (!grepl("\\.svg$", x, ignore.case = TRUE)) {
    dims <- image_dims_from_header(x)
  }
  if (!is.null(dims) && all(dims > 0)) {
    return(dims)
  }

  if (!requireNamespace("magick", quietly = TRUE)) {
    stop("package magick is required when using `guess_size` option.")
  }
  if (grepl("\\.svg$", x, ignore.case = TRUE)) {
    image <- magick::image_read_svg(x)
  } else {
    image <- magick::image_read(x)
  }
  dims <- attr(magick::image_data(image), "dim")[-1]
  c(width = dims[1], height = dims[2])
}
