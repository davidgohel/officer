#' @export
#' @importFrom utils zip
pack_folder <- function( folder, target ){
  curr_wd <- getwd()
  zip_dir <- folder
  setwd(zip_dir)
  zip(zipfile = target,
      files = list.files(all.files = TRUE, recursive = TRUE, include.dirs = TRUE))
  setwd(curr_wd)
  target
}

