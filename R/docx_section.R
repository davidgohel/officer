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
#' @note
#' This function is deprecated, use body_end_section_continuous,
#' body_end_section_landscape, body_end_section_portrait,
#' body_end_section_columns or body_end_section_columns_landscape
#' instead.
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
#'   slip_in_column_break(pos = "before") %>%
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
                             colwidths = c(1), space = .05, sep = FALSE, continuous = FALSE){# nocov start

  .Deprecated(msg = "body_end_section is deprecated. See ?sections for replacement functions.")

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
}# nocov end

#' @export
#' @rdname body_end_section
body_default_section <- function(x, landscape = FALSE,  margins = c(top = NA, bottom = NA, left = NA, right = NA)){# nocov start

  .Deprecated(msg = "body_default_section is deprecated. See ?sections for replacement functions.")

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
}# nocov end

#' @export
#' @rdname slip_in_column_break
break_column_before <- function( x ){ # nocov start
  .Deprecated(new = "slip_in_column_break")
  xml_elt <- paste0( wml_with_ns("w:r"), "<w:br w:type=\"column\"/>", "</w:r>")
  slip_in_xml(x = x, str = xml_elt, pos = "before")
} # nocov end

#' @export
#' @title add a column break
#' @description add a column break into a Word document. A column break
#' is used to add a break in a multi columns section in a Word
#' Document.
#' @param x an rdocx object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
slip_in_column_break <- function( x, pos = "before" ){
  xml_elt <- paste0( wml_with_ns("w:r"), "<w:br w:type=\"column\"/>", "</w:r>")
  slip_in_xml(x = x, str = xml_elt, pos = pos)
}




# new functions ----

#' @title sections
#'
#' @description Add sections in a Word document.
#'
#' @details
#' A section starts at the end of the previous section (or the beginning of
#' the document if no preceding section exists), and stops where the section is declared.
#' @param x an rdocx object
#' @param w,h width and height in inches of the section page. This will
#' be ignored if the default section (of the \code{reference_docx} file)
#' already has a width and a height.
#' @export
#' @rdname sections
#' @name sections
#' @examples
#' library(magrittr)
#'
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>%
#'   rep(5) %>% paste(collapse = "")
#' str2 <- "Aenean venenatis varius elit et fermentum vivamus vehicula. " %>%
#'   rep(5) %>% paste(collapse = "")
#'
#' my_doc <- read_docx()  %>%
#'   body_add_par(value = "Default section", style = "heading 1") %>%
#'   body_add_par(value = str1, style = "centered") %>%
#'   body_add_par(value = str2, style = "centered") %>%
#'
#'   body_end_section_continuous() %>%
#'   body_add_par(value = "Landscape section", style = "heading 1") %>%
#'   body_add_par(value = str1, style = "centered") %>%
#'   body_add_par(value = str2, style = "centered") %>%
#'   body_end_section_landscape() %>%
#'
#'   body_add_par(value = "Columns", style = "heading 1") %>%
#'   body_end_section_continuous() %>%
#'   body_add_par(value = str1, style = "centered") %>%
#'   body_add_par(value = str2, style = "centered") %>%
#'   slip_in_column_break() %>%
#'   body_add_par(value = str1, style = "centered") %>%
#'   body_end_section_columns(widths = c(2,2), sep = TRUE, space = 1) %>%
#'
#'   body_add_par(value = str1, style = "Normal") %>%
#'   body_add_par(value = str2, style = "Normal") %>%
#'   slip_in_column_break() %>%
#'   body_end_section_columns_landscape(widths = c(3,3), sep = TRUE, space = 1)
#'
#' print(my_doc, target = "section.docx")
body_end_section_continuous <- function( x ){
  str <- "<w:pPr><w:sectPr><w:officersection/><w:type w:val=\"continuous\"/></w:sectPr></w:pPr>"
  str <- paste0( wml_with_ns("w:p"), str, "</w:p>")
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @rdname sections
body_end_section_landscape <- function( x, w = 21 / 2.54, h = 29.7 / 2.54 ){
  w = w * 20 * 72
  h = h * 20 * 72
  pgsz_str <- "<w:pgSz w:orient=\"landscape\" w:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, h, w )
  str <- sprintf( "<w:pPr><w:sectPr><w:officersection/>%s</w:sectPr></w:pPr>", pgsz_str)
  str <- paste0( wml_with_ns("w:p"), str, "</w:p>")
  as_xml_document(str)
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @rdname sections
body_end_section_portrait <- function( x, w = 21 / 2.54, h = 29.7 / 2.54 ){
  w = w * 20 * 72
  h = h * 20 * 72
  pgsz_str <- "<w:pgSz w:orient=\"portrait\" w:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, w, h )
  str <- sprintf( "<w:pPr><w:sectPr><w:officersection/>%s</w:sectPr></w:pPr>", pgsz_str)
  str <- paste0( wml_with_ns("w:p"), str, "</w:p>")
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @param widths columns widths in inches. If 3 values, 3 columns
#' will be produced.
#' @param space space in inches between columns.
#' @param sep if TRUE a line is separating columns.
#' @rdname sections
body_end_section_columns <- function(x, widths = c(2.5,2.5), space = .25, sep = FALSE){

  widths <- widths * 20 * 72
  space <- space * 20 * 72

  columns_str_all_but_last <- sprintf("<w:col w:w=\"%.0f\" w:space=\"%.0f\"/>",
                                      widths[-length(widths)], space)
  columns_str_last <- sprintf("<w:col w:w=\"%.0f\"/>",
                              widths[length(widths)])
  columns_str <- c(columns_str_all_but_last, columns_str_last)

  if( length(widths) < 2 )
    stop("length of widths should be at least 2", call. = FALSE)

  columns_str <- sprintf("<w:cols w:num=\"%.0f\" w:sep=\"%.0f\" w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
                         length(widths), as.integer(sep), space, paste0(columns_str, collapse = "") )

  str <- paste0( "<w:pPr><w:sectPr><w:officersection/>",
                 "<w:type w:val=\"continuous\"/>",
                 columns_str, "</w:sectPr></w:pPr>")
  str <- paste0( wml_with_ns("w:p"), str, "</w:p>")
  body_add_xml(x, str = str, pos = "after")
}


#' @export
#' @rdname sections
body_end_section_columns_landscape <- function(x, widths = c(2.5,2.5), space = .25, sep = FALSE, w = 21 / 2.54, h = 29.7 / 2.54){

  widths <- widths * 20 * 72
  space <- space * 20 * 72

  columns_str_all_but_last <- sprintf("<w:col w:w=\"%.0f\" w:space=\"%.0f\"/>",
                                      widths[-length(widths)], space)
  columns_str_last <- sprintf("<w:col w:w=\"%.0f\"/>",
                              widths[length(widths)])
  columns_str <- c(columns_str_all_but_last, columns_str_last)

  if( length(widths) < 2 )
    stop("length of widths should be at least 2", call. = FALSE)

  columns_str <- sprintf("<w:cols w:num=\"%.0f\" w:sep=\"%.0f\" w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
                         length(widths), as.integer(sep), space, paste0(columns_str, collapse = "") )

  w = w * 20 * 72
  h = h * 20 * 72
  pgsz_str <- "<w:pgSz w:orient=\"landscape\" w:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, h, w )


  str <- paste0( "<w:pPr><w:sectPr><w:officersection/>",
                 pgsz_str,
                 columns_str, "</w:sectPr></w:pPr>")
  str <- paste0( wml_with_ns("w:p"), str, "</w:p>")
  body_add_xml(x, str = str, pos = "after")
}


# utils ----
process_sections <- function( x ){

  all_nodes <- xml_find_all(x$doc_obj$get(), "//w:sectPr[w:officersection]")
  main_sect <- xml_find_first(x$doc_obj$get(), "w:body/w:sectPr")

  for(node_id in seq_along(all_nodes) ){
    current_node <- as_xml_document(all_nodes[[node_id]])
    new_node <- as_xml_document(main_sect)

    # correct type ---
    type <- xml_child(current_node, "w:type")
    type_ref <- xml_child(new_node, "w:type")
    if( !inherits(type, "xml_missing") ){
      if( !inherits(type_ref, "xml_missing") )
        type_ref <- xml_replace(type_ref, type)
      else xml_add_child(new_node, type)
    }

    # correct cols ---
    cols <- xml_child(current_node, "w:cols")
    cols_ref <- xml_child(new_node, "w:cols")
    if( !inherits(cols, "xml_missing") ){
      if( !inherits(cols_ref, "xml_missing") )
        cols_ref <- xml_replace(cols_ref, cols)
      else xml_add_child(new_node, cols)
    }

    # correct pgSz ---
    pgSz <- xml_child(current_node, "w:pgSz")
    pgSz_ref <- xml_child(new_node, "w:pgSz")
    if( !inherits(pgSz, "xml_missing") ){

      if( !inherits(pgSz_ref, "xml_missing") ){
        xml_attr(pgSz_ref, "w:orient") <- xml_attr(pgSz, "orient")

        wref <- as.integer( xml_attr(pgSz_ref, "w") )
        href <- as.integer( xml_attr(pgSz_ref, "h") )

        if( xml_attr(pgSz, "orient") %in% "portrait" ){
          h <- ifelse( wref < href, href, wref )
          w <- ifelse( wref < href, wref, href )
        } else if( xml_attr(pgSz, "orient") %in% "landscape" ){
          w <- ifelse( wref < href, href, wref )
          h <- ifelse( wref < href, wref, href )
        }
        xml_attr(pgSz_ref, "w:w") <- w
        xml_attr(pgSz_ref, "w:h") <- h
      } else {
        xml_add_child(new_node, pgSz)
      }
    }

    node <- xml_replace(all_nodes[[node_id]], new_node)
  }
  x
}



section_dimensions <- function(node){
  section_obj <- as_list(node)

  landscape <- FALSE
  if( !is.null(attr(section_obj$pgSz, "orient")) && attr(section_obj$pgSz, "orient") == "landscape" ){
    landscape <- TRUE
  }

  h_ref <- as.integer(attr(section_obj$pgSz, "h"))
  w_ref <- as.integer(attr(section_obj$pgSz, "w"))

  mar_t <- as.integer(attr(section_obj$pgMar, "top"))
  mar_b <- as.integer(attr(section_obj$pgMar, "bottom"))
  mar_r <- as.integer(attr(section_obj$pgMar, "right"))
  mar_l <- as.integer(attr(section_obj$pgMar, "left"))
  mar_h <- as.integer(attr(section_obj$pgMar, "header"))
  mar_f <- as.integer(attr(section_obj$pgMar, "footer"))

  list( page = c("width" = w_ref, "height" = h_ref),
        landscape = landscape,
        margins = c(top = mar_t, bottom = mar_b,
                    left = mar_l, right = mar_r,
                    header = mar_h, footer = mar_f) )

}

