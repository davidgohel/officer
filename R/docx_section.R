#' @export
#' @title add section
#' @description add a section in a Word document. A section affects
#' preceding paragraphs or tables.
#'
#' @details
#' A section starts at the end of the previous section (or the beginning of
#' the document if no preceding section exists), and stops where the section is declared.
#' The function \code{body_end_section()} is reflecting that Word concept.
#' The function \code{body_default_section()} is only modifying the default section of
#' the document.
#' @importFrom xml2 xml_remove
#' @param x an rdocx object
#' @param landscape landscape orientation
#' @param margins a named vector of margin settings in inches, margins not set remain at their default setting
#' @param colwidths columns widths as percentages, summing to 1. If 3 values, 3 columns
#' will be produced.
#' @param space space in percent between columns.
#' @param sep if TRUE a line is separating columns.
#' @param continuous TRUE for a continuous section break.
#' @examples
#' library(magrittr)
#'
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>%
#'   rep(10) %>% paste(collapse = "")
#'
#' my_doc <- read_docx() %>%
#'   # add a paragraph
#'   body_add_par(value = str1, style = "Normal") %>%
#'   # add a continuous section
#'   body_end_section(continuous = TRUE) %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   # preceding paragraph is on a new column
#'   break_column_before() %>%
#'   # add a two columns continous section
#'   body_end_section(colwidths = c(.6, .4),
#'                    space = .05, sep = FALSE, continuous = TRUE) %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   # add a continuous section ... so far there is no break page
#'   body_end_section(continuous = TRUE) %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   body_default_section(landscape = TRUE, margins = c(top = 0.5, bottom = 0.5))
#'
#' print(my_doc, target = "section.docx")
body_end_section <- function(x, landscape = FALSE, margins = c(top = NA, bottom = NA, left = NA, right = NA),
                             colwidths = c(1), space = .05, sep = FALSE, continuous = FALSE){

  stopifnot(all.equal( sum(colwidths), 1 ) )

  if( landscape && continuous ){
    stop("using landscape=TRUE and continuous=TRUE is not possible as changing orientation require a new page.")
  }

  sdim <- x$sect_dim
  h_ref <- sdim$page["height"];w_ref <- sdim$page["width"]
  mar_t <- sdim$margins["top"];mar_b <- sdim$margins["bottom"]
  mar_r <- sdim$margins["right"];mar_l <- sdim$margins["left"]
  mar_h <- sdim$margins["header"];mar_f <- sdim$margins["footer"]

  if( !landscape ){
    h <- h_ref
    w <- w_ref
    mar_top <- mar_t
    mar_bottom <- mar_b
    mar_right <- mar_r
    mar_left <- mar_l
  } else {
    h <- w_ref
    w <- h_ref
    mar_top <- mar_r
    mar_bottom <- mar_l
    mar_right <- mar_t
    mar_left <- mar_b
  }

  if (!is.na(margins["top"]) & is.numeric(margins["top"])) { mar_top = margins["top"] * 20 * 72}
  if (!is.na(margins["bottom"]) & is.numeric(margins["bottom"])) { mar_bottom = margins["bottom"] * 20 * 72}
  if (!is.na(margins["left"]) & is.numeric(margins["left"])) { mar_left = margins["left"] * 20 * 72}
  if (!is.na(margins["right"]) & is.numeric(margins["right"])) { mar_right = margins["right"] * 20 * 72}

  pgsz_str <- "<w:pgSz %sw:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, ifelse( landscape, "w:orient=\"landscape\" ", ""), w, h )

  mar_str <- "<w:pgMar w:top=\"%.0f\" w:right=\"%.0f\" w:bottom=\"%.0f\" w:left=\"%.0f\" w:header=\"%.0f\" w:footer=\"%.0f\" w:gutter=\"0\"/>"
  mar_str <- sprintf(mar_str, mar_top, mar_right, mar_bottom, mar_left, mar_h, mar_f )

  width_ <- w - mar_right - mar_left
  column_values <- colwidths - space
  columns_str_all_but_last <- sprintf("<w:col w:w=\"%.0f\" w:space=\"%.0f\"/>",
                                      column_values[-length(column_values)] * width_,
                                      space * width_)
  columns_str_last <- sprintf("<w:col w:w=\"%.0f\"/>",
                              column_values[length(column_values)] * width_)
  columns_str <- c(columns_str_all_but_last, columns_str_last)
  if( length(colwidths) > 1 )
    columns_str <- sprintf("<w:cols w:num=\"%.0f\" w:sep=\"%.0f\" w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
                           length(colwidths), as.integer(sep), space * w, paste0(columns_str, collapse = "") )
  else columns_str <- sprintf("<w:cols w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
                              space * w, paste0(columns_str, collapse = "") )

  str <- paste0( wml_with_ns("w:p"),
                 "<w:pPr><w:sectPr>",
                 ifelse( continuous, "<w:type w:val=\"continuous\"/>", "" ),
                 pgsz_str, mar_str, columns_str, "</w:sectPr></w:pPr></w:p>")
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @rdname body_end_section
body_default_section <- function(x, landscape = FALSE,  margins = c(top = NA, bottom = NA, left = NA, right = NA)){

  sdim <- x$sect_dim
  h_ref <- sdim$page["height"];w_ref <- sdim$page["width"]
  mar_t <- sdim$margins["top"];mar_b <- sdim$margins["bottom"]
  mar_r <- sdim$margins["right"];mar_l <- sdim$margins["left"]
  mar_h <- sdim$margins["header"];mar_f <- sdim$margins["footer"]

  if( !landscape ){
    h <- h_ref
    w <- w_ref
    mar_top <- mar_t
    mar_bottom <- mar_b
    mar_right <- mar_r
    mar_left <- mar_l
  } else {
    h <- w_ref
    w <- h_ref
    mar_top <- mar_r
    mar_bottom <- mar_l
    mar_right <- mar_t
    mar_left <- mar_b
  }

  if (!is.na(margins["top"]) & is.numeric(margins["top"])) { mar_top = margins["top"] * 20 * 72}
  if (!is.na(margins["bottom"]) & is.numeric(margins["bottom"])) { mar_bottom = margins["bottom"] * 20 * 72}
  if (!is.na(margins["left"]) & is.numeric(margins["left"])) { mar_left = margins["left"] * 20 * 72}
  if (!is.na(margins["right"]) & is.numeric(margins["right"])) { mar_right = margins["right"] * 20 * 72}

  pgsz_str <- "<w:pgSz %sw:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, ifelse( landscape, "w:orient=\"landscape\" ", ""), w, h )

  mar_str <- "<w:pgMar w:top=\"%.0f\" w:right=\"%.0f\" w:bottom=\"%.0f\" w:left=\"%.0f\" w:header=\"%.0f\" w:footer=\"%.0f\" w:gutter=\"0\"/>"
  mar_str <- sprintf(mar_str, mar_top, mar_right, mar_bottom, mar_left, mar_h, mar_f )

  str <- paste0( wml_with_ns("w:sectPr"),
                 "<w:type w:val=\"continuous\"/>",
                 pgsz_str, mar_str, "</w:sectPr>")
  last_sect <- xml_find_first(x$doc_obj$get(), "/w:document/w:body/w:sectPr[last()]")
  xml_replace(last_sect, as_xml_document(str) )
  x
}

#' @export
#' @rdname body_end_section
break_column_before <- function( x ){
  xml_elt <- paste0( wml_with_ns("w:r"), "<w:br w:type=\"column\"/>", "</w:r>")
  slip_in_xml(x = x, str = xml_elt, pos = "before")
}

