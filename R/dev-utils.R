#' @title Write a ggplot Object to PNG File
#'
#' @description Renders a ggplot object to a PNG file using ragg for high-quality output.
#'
#' @param ggobj A ggplot object to render
#' @param width Numeric, width of the output image
#' @param height Numeric, height of the output image
#' @param res Numeric, resolution in DPI (default 200)
#' @param units Character, units for width and height ("in", "cm", "mm", "px") (default "in")
#' @param pointsize Integer, The default pointsize of the device in pt
#' @param scaling scaling factor to apply
#' @param path Character, output file path. If NULL, a temporary file is created (default NULL)
#'
#' @return Character, the path to the created PNG file
#'
#' @importFrom ragg agg_png
#' @importFrom grDevices dev.off
#' @examples
#' plot_in_png(
#'   code = {
#'     barplot(1:10)
#'   },
#'   width = 5,
#'   height = 4,
#'   res = 72,
#'   units = "in"
#' )
#' @export
#' @keywords internal
plot_in_png <- function(
  ggobj = NULL,
  code = NULL,
  width,
  height,
  res = 200,
  units = "in",
  pointsize = 11,
  scaling = 1,
  path = NULL
) {
  if (is.null(path)) {
    path <- tempfile(fileext = ".png")
  }

  if (!is.null(ggobj)) {
    code <- str2lang("print(ggobj)")
  }

  agg_png(
    filename = path,
    width = width,
    height = height,
    units = units,
    background = "transparent",
    res = res,
    pointsize = pointsize,
    scaling = scaling
  )
  tryCatch(
    {
      eval(code)
    },
    finally = dev.off()
  )
  path
}

#' @importFrom openssl base64_encode
#' @export
#' @title Encode Character Vector to Base64
#' @description
#' Encodes one or more elements of a character vector into Base64 format.
#' @param x A character vector. NA values are preserved.
#' @return A character vector of Base64-encoded strings, same length as `x`.
#' @examples
#' as_base64(letters)
#' as_base64(c("hello", NA, "world"))
#' @keywords internal
as_base64 <- function(x) {
  if (!is.character(x)) {
    stop("'x' must be a character vector.")
  }
  z <- vapply(
    x,
    function(elem) {
      if (is.na(elem)) {
        NA_character_
      } else {
        base64_encode(charToRaw(elem))
      }
    },
    NA_character_
  )
  unname(z)
}

#' @importFrom openssl base64_decode
#' @export
#' @title Decode Base64 Vector to Character
#' @description
#' Decodes one or more Base64-encoded elements back into their original character form.
#' @param x A character vector of Base64-encoded strings. NA values are preserved.
#' @return A character vector of decoded strings, same length as `x`.
#' @examples
#' z <- as_base64(c("hello", "world"))
#' from_base64(z)
#' @keywords internal
from_base64 <- function(x) {
  if (!is.character(x)) {
    stop("'x' must be a character vector of Base64 strings.")
  }
  z <- vapply(
    x,
    function(elem) {
      if (is.na(elem)) {
        NA_character_
      } else {
        raw <- tryCatch(
          {
            b64_str <- base64_decode(elem)
            if(all(nchar(b64_str) < 1)) {
              stop("empty result")
            }
            b64_str
          },
          error = function(e) {
            stop("Failed to decode Base64 element: '", elem, "'.", call. = FALSE)
          }
        )
        rawToChar(raw)
      }
    },
    NA_character_
  )
  unname(z)
}


#' Convert Data URIs to PNG Files
#'
#' @description Decodes base64-encoded data URIs and writes them to PNG files.
#'
#' @param data_uri Character, a data URI character vector starting with "data:image/png;base64,"
#' @param output_files Character, paths to the output PNG files
#'
#' @return Character, the paths to the created PNG files
#'
#' @importFrom openssl base64_decode
#' @examples
#' rlogo <- file.path(R.home("doc"), "html", "logo.jpg")
#' base64_str <- image_to_base64(rlogo)
#' base64_to_image(
#'   data_uri = base64_str,
#'   output_files = tempfile(fileext = ".jpeg")
#' )
#' @export
#' @keywords internal
base64_to_image <- function(data_uri, output_files) {

  for(i in seq_along(data_uri)) {
    base64_part <- sub("^data:image/[^;]+;base64,", "", data_uri[[i]])
    raw_data <- base64_decode(base64_part)
    writeBin(raw_data, output_files[[i]])
  }

  output_files
}

mime_type <- function(paths) {
  result <- character(length(paths))
  pattern <- "\\.(png|gif|jpg|jpeg|svg|tiff|pdf|webp)$"
  m <- regexpr(pattern = pattern, text = paths)
  result[attr(m, "match.length") > -1] <- regmatches(paths, m)
  result <- gsub("\\.jpg", ".jpeg", result)
  result <- gsub("\\.svg", ".svg+xml", result)
  result <- gsub("^\\.{1}", "", result)
  prefix <- ifelse(result %in% "pdf", "application", "image")
  result <- paste(prefix, result, sep = "/")
  result[attr(m, "match.length") < 0] <- NA_character_
  result
}

#' @importFrom openssl base64_encode
#' @export
#' @title Images to base64
#' @description encodes images into base64 strings.
#' @param filepaths file names.
#' @keywords internal
#' @examples
#' rlogo <- file.path(R.home("doc"), "html", "logo.jpg")
#' base64_str <- image_to_base64(rlogo)
image_to_base64 <- function(filepaths) {

  mimes <- mime_type(paths = filepaths)
  if (any(is.na(mimes))) {
    cli::cli_abort(
      paste0(
        "Unknown image(s) format: ",
        cli::ansi_collapse(
          basename(filepaths)[is.na(mimes)],
          trunc = 5
        )
      )
    )
  }

  if (any(!file.exists(filepaths))) {
    cli::cli_abort(
      paste0(
        "File(s) not found: ",
        cli::ansi_collapse(
          basename(filepaths)[!file.exists(filepaths)],
          trunc = 5
        )
      )
    )
  }

  base64_lst <- mapply(
    FUN = function(filepath, mime) {
      dat <- readBin(
        filepath,
        what = "raw",
        size = 1,
        endian = "little",
        n = file.info(filepath)$size
      )
      base64_str <- base64_encode(bin = dat)
      base64_str <- sprintf("data:%s;base64,%s", mime, base64_str)
      base64_str
    },
    filepath = filepaths,
    mime = mimes,
    SIMPLIFY = FALSE
  )
  unname(unlist(base64_lst))
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
  `&` = "&amp;",
  `<` = "&lt;",
  `>` = "&gt;",
  `'` = "&#39;",
  `"` = "&quot;",
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
