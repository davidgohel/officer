#' @export
#' @importFrom R.utils getAbsolutePath
#' @importFrom utils zip
#' @title compress a folder
#' @description compress a folder to a target file. The
#' function returns the complete path to target file.
#' @details
#' The function is using \link[utils]{zip}, it needs a zip program.
#' @param folder folder to compress
#' @param target path of the archive to create
pack_folder <- function( folder, target ){

  target <- getAbsolutePath(target)
  folder <- getAbsolutePath(folder)
  curr_wd <- getwd()
  zip_dir <- folder
  setwd(zip_dir)
  zip(zipfile = target, flags = "-qr9X",
      files = list.files(all.files = TRUE, recursive = TRUE))
  setwd(curr_wd)

  if( !file.exists(target) ){
    if( Sys.getenv("R_ZIPCMD") == "")
      stop('A zip application must be available to R. You can set its path with Sys.setenv("R_ZIPCMD" = "path/to/zip")', call. = FALSE)
    msg <- sprintf("could not zip package %s in file %s", shQuote(zip_dir), shQuote(target))
    stop(msg, call. = FALSE)
  }

  target
}

#' @export
#' @importFrom utils unzip
#' @importFrom R.utils getAbsolutePath
#' @title Extract files from a zip file
#' @description Extract files from a zip file to a folder. The
#' function returns the complete path to destination folder.
#' @details
#' The function is using \link[utils]{unzip}, it needs an unzip program.
#' @param file path of the archive to unzip
#' @param folder folder to create
unpack_folder <- function( file, folder ){

  file <- getAbsolutePath(file)
  folder <- getAbsolutePath(folder)
  stopifnot(file.exists(file))

  unlink(folder, recursive = TRUE, force = TRUE)

  unzip( zipfile = file, exdir = folder )

  folder
}

#' @export
#' @title test zip function
#' @description test wether zip can produce a zip file.
#' @examples
#' has_zip()
has_zip <- function(){
  ifile <- tempfile(fileext = ".txt")
  cat("hi", file = ifile)
  ofile <- tempfile(fileext = ".zip")
  try(zip(zipfile = ofile, flags = "-qr9X", files = ifile ), silent = TRUE)
  file.exists(ofile)
}
