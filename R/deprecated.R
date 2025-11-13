#' @export
#' @title add images into an rdocx object
#' @description reference images into a Word document.
#' This function is now useless as the processing of images
#' is automated when using [print.rdocx()].
#'
#' @param x an rdocx object
#' @param src a vector of character containing image filenames.
#' @keywords internal
docx_reference_img <- function(x, src) {
  .Deprecated("",
    package = "officer",
    msg = paste(
      "The `docx_reference_img()` function is no longer useful.\n",
      "Please remove calls to `docx_reference_img()` as this function will be removed in a future version."
    )
  )
  x
}
