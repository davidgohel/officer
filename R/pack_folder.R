#' @export
#' @importFrom utils zip
#' @importFrom R.utils getAbsolutePath
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
  target
}

