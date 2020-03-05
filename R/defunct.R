#' @name officer-defunct
#' @title Defunct Functions in Package officer
NULL

#' @export
#' @rdname officer-defunct
#' @details `ph_from_xml()` is replaced by `ph_with.xml_document`.
#' @param ... unused arguments
ph_from_xml <- function(...){
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_from_xml_at()` is replaced by `ph_with.xml_document`.
ph_from_xml_at <- function(...){
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_with_table()` is replaced by `ph_with.xml_document`.
ph_with_table <- function(...){
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_with_img()` is replaced by `ph_with.xml_document`.
ph_with_img <- function(...){
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_with_gg()` is replaced by `ph_with.xml_document`.
ph_with_gg <- function(...){
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_with_ul()` is replaced by `ph_with.xml_document`.
ph_with_ul <- function(...) {
  .Defunct("ph_with")
}

#' @export
#' @rdname officer-defunct
#' @details `ph_with_table_at()` is replaced by `ph_with.xml_document`.
ph_with_table_at <- function(...){
  .Defunct(new = "ph_with")
}


#' @export
#' @rdname officer-defunct
#' @details `ph_with_fpars_at()` is replaced by `ph_with.xml_document`.
ph_with_fpars_at <- function(...){
  .Defunct(new = "ph_with")
}


#' @export
#' @rdname officer-defunct
#' @details `body_end_section()` is replaced by function `body_end_section_*`.
body_end_section <- function(...){
  .Defunct(new = "body_end_section_*")
}

#' @export
#' @rdname officer-defunct
#' @details `body_default_section()` is replaced by function `body_end_section_*`.
body_default_section <- function(...){
  .Defunct(new = "body_end_section_*")
}

#' @export
#' @rdname officer-defunct
#' @details `break_column_before()` is replaced by function `slip_in_column_break`.
break_column_before <- function(...){
  .Defunct(new = "slip_in_column_break")
}

