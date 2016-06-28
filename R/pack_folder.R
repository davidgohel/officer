#' @export
#' @importFrom utils zip
#' @importFrom R.utils getAbsolutePath
pack_folder <- function( folder, target ){
  target <- getAbsolutePath(target)
  curr_wd <- getwd()
  zip_dir <- folder
  setwd(zip_dir)
  zip(zipfile = target, flags = "-qr9X",
      files = list.files(all.files = TRUE, recursive = TRUE))
  setwd(curr_wd)
  target
}

