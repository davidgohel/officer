#' @export
#' @importFrom zip zip
#' @title compress a folder
#' @description compress a folder to a target file. The
#' function returns the complete path to target file.
#' @param folder folder to compress
#' @param target path of the archive to create
#' @keywords internal
pack_folder <- function(folder, target) {
  target <- absolute_path(target)
  dir_fi <- dirname(target)

  if (!file.exists(dir_fi)) {
    stop("directory ", shQuote(dir_fi), " does not exist.", call. = FALSE)
  } else if (file.access(dir_fi) < 0) {
    stop("can not write to directory ", shQuote(dir_fi), call. = FALSE)
  } else if (file.exists(target) && file.access(target) < 0) {
    stop(shQuote(target), " already exists and is not writable", call. = FALSE)
  } else if (!file.exists(target)) {
    old_warn <- getOption("warn")
    on.exit(options(warn = old_warn), add = TRUE)
    options(warn = -1)
    x <- tryCatch(
      {
        cat("", file = target)
        TRUE
      },
      error = function(e) FALSE,
      finally = unlink(target, force = TRUE)
    )
    if (!x) {
      stop(
        shQuote(target),
        " cannot be written, please check your permissions.",
        call. = FALSE
      )
    }
  }

  # All files, hidden files at the root of the folder excepted, sorted in an
  # order that does not depend on the locale.
  files <- list.files(path = folder, recursive = TRUE, all.files = TRUE)
  files <- files[!startsWith(sub("/.*$", "", files), ".")]
  files <- sort(files, method = "radix")

  tryCatch(
    zip::zipr(
      zipfile = target,
      files = files,
      recurse = FALSE,
      include_directories = FALSE,
      root = folder,
      mode = "mirror"
    ),
    error = function(e) {
      stop("Could not write ", shQuote(target), " [", e$message, "]")
    }
  )

  target
}

# Write the same date in all entries of a zip file, the one Office writes:
# 1980-01-01 00:00. The date that zip stores is the date of the file converted
# to local time by the C library, which does not follow the time zone of R on
# Windows, so files written at another time or in another time zone differ.
# The date and time fields are rewritten in the central directory and in the
# local header of each entry. Zip64 archives (more than 4 GB or 65535 entries)
# are left as they are.
fix_zip_dates <- function(path) {
  size <- file.size(path)
  bytes <- readBin(path, what = "raw", n = size)

  u16 <- function(i) as.integer(bytes[i]) + 256L * as.integer(bytes[i + 1L])
  u32 <- function(i) sum(as.numeric(bytes[i:(i + 3L)]) * 256^(0:3))
  is_sig <- function(i, sig) identical(bytes[i:(i + 3L)], as.raw(sig))

  # the end of central directory record is in the last 65557 bytes
  eocd <- NA_integer_
  for (i in seq.int(size - 21L, max(1L, size - 65556L))) {
    if (is_sig(i, c(0x50, 0x4b, 0x05, 0x06))) {
      eocd <- i
      break
    }
  }
  if (is.na(eocd)) {
    return(invisible(path))
  }
  n_entries <- u16(eocd + 10L)
  cd_offset <- u32(eocd + 16L)
  if (n_entries == 0xFFFF || cd_offset == 0xFFFFFFFF) {
    return(invisible(path))
  }

  dos_time <- as.raw(c(0x00, 0x00))
  dos_date <- as.raw(c(0x21, 0x00)) # 1980-01-01
  pos <- cd_offset + 1
  for (k in seq_len(n_entries)) {
    if (!is_sig(pos, c(0x50, 0x4b, 0x01, 0x02))) {
      cli::cli_abort("Unexpected structure in zip file {.file {path}}.")
    }
    bytes[pos + 12:13] <- dos_time
    bytes[pos + 14:15] <- dos_date
    local <- u32(pos + 42L) + 1
    if (!is_sig(local, c(0x50, 0x4b, 0x03, 0x04))) {
      cli::cli_abort("Unexpected structure in zip file {.file {path}}.")
    }
    bytes[local + 10:11] <- dos_time
    bytes[local + 12:13] <- dos_date
    pos <- pos + 46 + u16(pos + 28L) + u16(pos + 30L) + u16(pos + 32L)
  }

  writeBin(bytes, path)
  invisible(path)
}

#' @export
#' @importFrom zip unzip
#' @title Extract files from a zip file
#' @description Extract files from a zip file to a folder. The
#' function returns the complete path to destination folder.
#' @param file path of the archive to unzip
#' @param folder folder to create
#' @keywords internal
unpack_folder <- function(file, folder) {
  stopifnot(file.exists(file))

  file_type <- gsub("(.*)(\\.[a-zA-Z0-9]+)$", "\\2", file)

  # force deletion if already existing
  unlink(folder, recursive = TRUE, force = TRUE)

  if (l10n_info()$`UTF-8`) {
    zip::unzip(zipfile = file, exdir = folder)
  } else {
    officer_wd_folder <- getOption("officer_wd_folder")
    if (is.null(officer_wd_folder) || !dir.exists(officer_wd_folder)) {
      officer_wd_folder <- tempdir()
    }
    # unable to unzip a file with accent when on windows
    newfile <- tempfile(tmpdir = officer_wd_folder, fileext = file_type)
    file.copy(from = file, to = newfile)
    zip::unzip(zipfile = newfile, exdir = folder)
    unlink(newfile, force = TRUE)
  }

  absolute_path(folder)
}

absolute_path <- function(x) {
  if (length(x) != 1L) {
    stop("'x' must be a single character string")
  }
  epath <- path.expand(x)

  if (file.exists(epath)) {
    epath <- normalizePath(epath, "/", mustWork = TRUE)
  } else {
    if (!dir.exists(dirname(epath))) {
      stop("directory of ", x, " does not exist.", call. = FALSE)
    }
    cat("", file = epath)
    epath <- normalizePath(epath, "/", mustWork = TRUE)
    unlink(epath)
  }
  epath
}
