#' @export
#' @importFrom utils compareVersion packageVersion
#' @import zip
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

  if( compareVersion(as.character(packageVersion("zip")), "1.0.0") > 0 ){
    ## replacement when zip will be greater than 1.0.0
    # zip::zipr(zipfile = target,
    #           files = list.files(path = folder, all.files = FALSE, full.names = TRUE),
    #           recurse = TRUE)
    call_ <- call("zipr", zipfile = target,
                  files = list.files(path = folder, all.files = FALSE, full.names = TRUE),
                  recurse = TRUE)
    eval(call_)
    return(target)
  }

  curr_wd <- getwd()
  setwd(folder)
  if( .Platform$OS.type %in% "windows")
    target <- enc2native(target)
  tryCatch(
    zip::zip(zipfile = target,
        files = list.files(all.files = TRUE, recursive = TRUE))
    , error = function(e) {
      stop("Could not write ", shQuote(target), " [", e$message, "]")
    }
    , finally = {
      setwd(curr_wd)
    })

  target
}

#' @export
#' @importFrom utils unzip
#' @title Extract files from a zip file
#' @description Extract files from a zip file to a folder. The
#' function returns the complete path to destination folder.
#' @details
#' The function is using \link[utils]{unzip}, it needs an unzip program.
#' @param file path of the archive to unzip
#' @param folder folder to create
unpack_folder <- function( file, folder ){

  stopifnot(file.exists(file))

  unlink(folder, recursive = TRUE, force = TRUE)

  unzip( zipfile = file, exdir = folder )

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

