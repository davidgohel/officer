#' @export
#' @importFrom zip zip
#' @title compress a folder
#' @description compress a folder to a target file. The
#' function returns the complete path to target file.
#' @param folder folder to compress
#' @param target path of the archive to create
pack_folder <- function( folder, target ){

  target <- absolute_path(target)
  dir_fi <- dirname(target)

  if( !file.exists(dir_fi) ){
    stop("directory ", shQuote(dir_fi), " does not exist.", call. = FALSE)
  } else if( file.access(dir_fi) < 0 ){
    stop("can not write to directory ", shQuote(dir_fi), call. = FALSE)
  } else if( file.exists(target) && file.access(target) < 0 ){
    stop(shQuote(target), " already exists and is not writable", call. = FALSE)
  } else if( !file.exists(target) ){
    old_warn <- getOption("warn")
    options(warn = -1)
    x <- tryCatch({cat("", file = target);TRUE}, error = function(e) FALSE, finally = unlink(target, force = TRUE) )
    options(warn = old_warn)
    if( !x )
      stop(shQuote(target), " cannot be written, please check your permissions.", call. = FALSE)
  }

  curr_wd <- getwd()
  setwd(folder)
  tryCatch(
    zip::zipr(zipfile = target, include_directories = FALSE,
                files = list.files(path = ".", all.files = FALSE), recurse = TRUE)
    , error = function(e) {
      stop("Could not write ", shQuote(target), " [", e$message, "]")
    },
    finally = {
      setwd(curr_wd)
    })

  target
}

#' @export
#' @importFrom zip unzip
#' @title Extract files from a zip file
#' @description Extract files from a zip file to a folder. The
#' function returns the complete path to destination folder.
#' @param file path of the archive to unzip
#' @param folder folder to create
unpack_folder <- function( file, folder ){

  stopifnot(file.exists(file))

  file_type <- gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", file)

  # force deletion if already existing
  unlink(folder, recursive = TRUE, force = TRUE)

  if( l10n_info()$`UTF-8` ){
    zip::unzip( zipfile = file, exdir = folder )
  } else {
    # unable to unzip a file with accent when on windows
    newfile <- tempfile(fileext = file_type)
    file.copy(from = file, to = newfile)
    zip::unzip( zipfile = newfile, exdir = folder )
    unlink(newfile, force = TRUE)
  }

  absolute_path(folder)
}

absolute_path <- function(x){

  if (length(x) != 1L)
    stop("'x' must be a single character string")
  epath <- path.expand(x)

  if( file.exists(epath)){
    epath <- normalizePath(epath, "/", mustWork = TRUE)
  } else {
    if( !dir.exists(dirname(epath)) ){
      stop("directory of ", x, " does not exist.", call. = FALSE)
    }
    cat("", file = epath)
    epath <- normalizePath(epath, "/", mustWork = TRUE)
    unlink(epath)
  }
  epath
}

