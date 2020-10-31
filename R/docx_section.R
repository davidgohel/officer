#' @export
#' @title add a column break
#' @description add a column break into a Word document. A column break
#' is used to add a break in a multi columns section in a Word
#' Document.
#' @param x an rdocx object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
slip_in_column_break <- function( x, pos = "before" ){
  xml_elt <- paste0( wr_ns_yes, "<w:br w:type=\"column\"/>", "</w:r>")
  slip_in_xml(x = x, str = xml_elt, pos = pos)
}



#' @export
#' @title add continuous section
#' @description Section break starts the new section on the same page. This
#' type of section break is often used to change the number of columns
#' without starting a new page.
#' @param x an rdocx object
#' @examples
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
#' str1 <- rep(str1, 5)
#' str1 <- paste(str1, collapse = " ")
#' str2 <- "Aenean venenatis varius elit et fermentum vivamus vehicula."
#' str2 <- rep(str2, 5)
#' str2 <- paste(str2, collapse = " ")
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, value = "Default section", style = "heading 1")
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_add_par(doc_1, value = str2, style = "Normal")
#' doc_1 <- body_end_section_continuous(doc_1)
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for Word sections
body_end_section_continuous <- function( x ){
  bs <- block_section(prop_section(type = "continuous"))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @title add landscape section
#' @description A section with landscape orientation is added to the document.
#' @param x an rdocx object
#' @param w,h page width, page height (in inches)
#' @examples
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
#' str1 <- rep(str1, 5)
#' str1 <- paste(str1, collapse = " ")
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_end_section_landscape(doc_1)
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for Word sections
body_end_section_landscape <- function( x, w = 21 / 2.54, h = 29.7 / 2.54){
  bs <- block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "landscape"),
    type = "oddPage"))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @title add portrait section
#' @description A section with portrait orientation is added to the document.
#' @param x an rdocx object
#' @param w,h page width, page height (in inches)
#' @examples
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
#' str1 <- rep(str1, 5)
#' str1 <- paste(str1, collapse = " ")
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_end_section_portrait(doc_1)
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for Word sections
body_end_section_portrait <- function( x, w = 21 / 2.54, h = 29.7 / 2.54){
  bs <- block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "portrait"),
    type = "oddPage"))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @title add multi columns section
#' @description A section with multiple columns is added to the document.
#' @param x an rdocx object
#' @inheritParams section_columns
#' @examples
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
#' str1 <- rep(str1, 5)
#' str1 <- paste(str1, collapse = " ")
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_end_section_columns(doc_1)
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for Word sections
body_end_section_columns <- function(x, widths = c(2.5,2.5), space = .25, sep = FALSE){

  bs <- block_section(prop_section(
    section_columns = section_columns(widths = widths, space = space, sep = sep),
    type = NULL))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}


#' @export
#' @title add multi columns section within landscape orientation
#' @description A landscape section with multiple columns is added to the document.
#' @param x an rdocx object
#' @inheritParams section_columns
#' @param w,h page width, page height (in inches)
#' @examples
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
#' str1 <- rep(str1, 5)
#' str1 <- paste(str1, collapse = " ")
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- slip_in_column_break(doc_1, pos = "after")
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_end_section_columns_landscape(doc_1, widths = c(6, 2))
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for Word sections
body_end_section_columns_landscape <- function(x, widths = c(2.5,2.5), space = .25,
                                               sep = FALSE, w = 21 / 2.54, h = 29.7 / 2.54){

  bs <- block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "landscape"),
    section_columns = section_columns(widths = widths, space = space, sep = sep),
    type = "oddPage"))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}
#' @export
#' @title add any section
#' @description Add a section to the document. You can
#' define any section with a [block_section] object. All other
#' `body_end_section_*` are specialized, this one is highly flexible
#' but it's up to the user to define the section properties.
#' @param x an rdocx object
#' @param value a [block_section] object
#' @examples
#' library(officer)
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
#' str1 <- rep(str1, 20)
#' str1 <- paste(str1, collapse = " ")
#'
#' ps <- prop_section(
#'   page_size = page_size(orient = "landscape"),
#'   page_margins = page_mar(top = 2),
#'   type = "continuous"
#' )
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#'
#' doc_1 <- body_end_block_section(doc_1, block_section(ps))
#'
#' doc_1 <- body_add_par(doc_1, value = str1, style = "centered")
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for Word sections
#' @section Illustrations:
#'
#' \if{html}{\figure{body_end_block_section_doc_1.png}{options: width=80\%}}
body_end_block_section <- function( x, value ){

  stopifnot(inherits(value, "block_section"))
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml(x = x, str = xml_elt, pos = "after")

  x
}

#' @export
#' @title Define Default Section
#' @description Define default section of the document. You can
#' define section propeerties (page size, orientation, ...) with a [prop_section] object.
#' @param x an rdocx object
#' @param value a [prop_section] object
#' @family functions for Word sections
#' @examples
#' default_sect_properties <- prop_section(
#'     page_size = page_size(orient = "landscape"), type = "continuous",
#'     page_margins = page_mar(bottom = .75, top = 1.5, right = 2, left = 2)
#'   )
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_table(doc_1, value = mtcars[1:10,], style = "table_template")
#' doc_1 <- body_add_par(doc_1, value = paste(rep(letters, 40), collapse = " "))
#' doc_1 <- body_set_default_section(doc_1, default_sect_properties)
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @section Illustrations:
#'
#' \if{html}{\figure{body_set_default_section_doc_1.png}{options: width=80\%}}
body_set_default_section <- function( x, value ){

  stopifnot(inherits(value, "prop_section"))
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  z <- xml_find_first(docx_body_xml(x), "w:body/w:sectPr")
  xml_replace(z, as_xml_document(xml_elt))

  x
}

# utils ----
process_sections <- function( x ){

  all_nodes <- xml_find_all(x$doc_obj$get(), "//w:pPr/w:sectPr")
  main_sect <- xml_find_first(x$doc_obj$get(), "w:body/w:sectPr")

  footer_refs <- xml_find_all(x$doc_obj$get(), "w:body/w:sectPr/w:footerReference")
  header_refs <- xml_find_all(x$doc_obj$get(), "w:body/w:sectPr/w:headerReference")

  for(node_id in seq_along(all_nodes) ){
    current_node <- all_nodes[[node_id]]

    header_ref <- xml_child(current_node, "w:headerReference")
    footer_ref <- xml_child(current_node, "w:footerReference")

    if(inherits(footer_ref, "xml_missing") && length(footer_refs) > 0 ){
      for(l in rev(seq_along(footer_refs))){
        xml_add_child(current_node, footer_refs[[l]], .where = 1)
      }
    }
    if(inherits(header_ref, "xml_missing") && length(header_refs) > 0 ){
      for(l in rev(seq_along(header_refs))){
        xml_add_child(current_node, header_refs[[l]], .where = 1)
      }
    }

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

