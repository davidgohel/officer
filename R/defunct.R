#' @name officer-defunct
#' @title Defunct Functions in Package officer
NULL


#' @export
#' @rdname officer-defunct
#' @details `ph_add_text()` is replaced by `ph_with(value = fpar(...))`.
#' @param ... unused arguments
ph_add_text <- function(...){
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_add_par()` is replaced by `ph_with(value = fpar(...))`.
#' @param ... unused arguments
ph_add_par <- function(...){
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_add_fpar()` is replaced by `ph_with(value = fpar(...))`.
#' @param ... unused arguments
ph_add_fpar <- function(...){
  .Defunct("ph_with")
}


#' @export
#' @rdname officer-defunct
#' @details `slip_in_seqfield()` is replaced by `run_word_field()`.
#' @param ... unused arguments
slip_in_seqfield <- function(...){
  .Defunct("run_word_field")
}


#' @export
#' @rdname officer-defunct
#' @details `slip_in_img()` is replaced by `external_img()`.
#' @param ... unused arguments
slip_in_img <- function(...){
  .Defunct("external_img")
}
