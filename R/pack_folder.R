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
  zz <- zip(zipfile = target, flags = "-qr9X",
      files = list.files(all.files = TRUE, recursive = TRUE))
  setwd(curr_wd)
  if(zz != 0){
    stop('could not zip document package. A zip application must be available to R.
         Set its path with Sys.setenv("R_ZIPCMD" = "path/to/zip")', call. = FALSE)
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
