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

# Give all files of a folder the same date before they are zipped, the one
# Office writes: the date stored in a zip file is the local time, so that
# files written at another time or in another time zone would differ.
reset_file_times <- function(folder) {
  files <- list.files(path = folder, recursive = TRUE, all.files = TRUE,
                      full.names = TRUE)
  Sys.setFileTime(files, as.POSIXct("1980-01-01 00:00:00"))
  invisible(folder)
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
