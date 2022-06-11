#' @name officer-defunct
#' @title Defunct Functions in Package officer
#' @export
#' @rdname officer-defunct
#' @details `slip_in_seqfield()` is replaced by `run_word_field()`.
#' @param ... unused arguments
slip_in_seqfield <- function(...){
  .Defunct("run_word_field")
}

#' @export
#' @rdname officer-defunct
#' @details `slip_in_column_break()` is replaced by `run_columnbreak()`.
#' @param ... unused arguments
slip_in_column_break <- function(...){
  .Defunct("run_columnbreak")
}

#' @export
#' @rdname officer-defunct
#' @details `slip_in_xml()` is replaced by `fpar()`.
#' @param ... unused arguments
slip_in_xml <- function(...){
  .Defunct("fpar()")
}

#' @export
#' @rdname officer-defunct
#' @details `slip_in_text()` is replaced by `fpar()`.
#' @param ... unused arguments
slip_in_text <- function(...){
  .Defunct("fpar()")
}


