#' @export
#' @title Add continuous section
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
body_end_section_continuous <- function(x) {
  bs <- block_section(prop_section(type = "continuous"))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @title Add landscape section
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
body_end_section_landscape <- function(x, w = 21 / 2.54, h = 29.7 / 2.54) {
  bs <- block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "landscape"),
    type = "oddPage"
  ))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @title Add portrait section
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
body_end_section_portrait <- function(x, w = 21 / 2.54, h = 29.7 / 2.54) {
  bs <- block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "portrait"),
    type = "oddPage"
  ))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}

#' @export
#' @title Add multi columns section
#' @description A section with multiple columns is added to the
#' document.
#'
#' You may prefer to use [body_end_block_section()] that is
#' more flexible.
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
body_end_section_columns <- function(x, widths = c(2.5, 2.5), space = .25, sep = FALSE) {
  bs <- block_section(prop_section(
    section_columns = section_columns(widths = widths, space = space, sep = sep),
    type = "continuous"
  ))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}


#' @export
#' @title Add a landscape multi columns section
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
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' doc_1 <- body_end_section_columns_landscape(doc_1, widths = c(6, 2))
#' doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for Word sections
body_end_section_columns_landscape <- function(x, widths = c(2.5, 2.5), space = .25,
                                               sep = FALSE, w = 21 / 2.54, h = 29.7 / 2.54) {
  bs <- block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "landscape"),
    section_columns = section_columns(widths = widths, space = space, sep = sep),
    type = "oddPage"
  ))
  str <- to_wml(bs, add_ns = TRUE)
  body_add_xml(x, str = str, pos = "after")
}
#' @export
#' @title Add any section
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
body_end_block_section <- function(x, value) {
  stopifnot(inherits(value, "block_section"))
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml(x = x, str = xml_elt, pos = "after")
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
#'   page_size = page_size(orient = "landscape"), type = "continuous",
#'   page_margins = page_mar(bottom = .75, top = 1.5, right = 2, left = 2)
#' )
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_table(doc_1, value = mtcars[1:10, ], style = "table_template")
#' doc_1 <- body_add_par(doc_1, value = paste(rep(letters, 40), collapse = " "))
#' doc_1 <- body_set_default_section(doc_1, default_sect_properties)
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @section Illustrations:
#'
#' \if{html}{\figure{body_set_default_section_doc_1.png}{options: width=80\%}}
body_set_default_section <- function(x, value) {
  stopifnot(inherits(value, "prop_section"))
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  z <- xml_find_first(docx_body_xml(x), "w:body/w:sectPr")
  xml_replace(z, as_xml_document(xml_elt))

  x
}


# utils ----
process_sections <- function(x) {
  all_nodes <- xml_find_all(x$doc_obj$get(), "//w:pPr/w:sectPr")

  default_pgMar <- xml_find_first(x$doc_obj$get(), "w:body/w:sectPr/w:pgMar")
  default_pgSz <- xml_find_first(x$doc_obj$get(), "w:body/w:sectPr/w:pgSz")
  sect_dim <- section_dimensions(xml_find_first(x$doc_obj$get(), "w:body/w:sectPr"))

  node_def_sec <- xml_find_first(x$doc_obj$get(), "w:body/w:sectPr")

  # if w:type not there, each section is on a new page if not continuous
  if (inherits(xml_child(node_def_sec, "w:type"), "xml_missing")) {
    node_type <- as_xml_document("<w:type xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:val=\"continuous\"/>")
    xml_add_child(node_def_sec, node_type)
  }

  x <- officer_section_fortify(node_def_sec, x)

  for (node_id in seq_along(all_nodes)) {
    current_node <- all_nodes[[node_id]]

    x <- officer_section_fortify(current_node, x)

    current_pgSz <- xml_child(current_node, "w:pgSz")
    if (inherits(current_pgSz, "xml_missing") && length(default_pgSz) > 0) {
      xml_add_child(current_node, default_pgSz)
    } else if (!inherits(current_pgSz, "xml_missing")) {
      fill_if_missing(current_pgSz, "h", sect_dim$page["height"])
      fill_if_missing(current_pgSz, "w", sect_dim$page["width"])
      if (length(sect_dim$landscape) && sect_dim$landscape) {
        fill_if_missing(current_pgSz, "orient", "landscape")
      }
    }

    current_pgMar <- xml_child(current_node, "w:pgMar")
    if (inherits(current_pgMar, "xml_missing") && length(default_pgMar) > 0) {
      xml_add_child(current_node, default_pgMar)
    } else if (!inherits(current_pgMar, "xml_missing")) {
      fill_if_missing(current_pgMar, "header", sect_dim$margins["header"])
      fill_if_missing(current_pgMar, "footer", sect_dim$margins["footer"])
      fill_if_missing(current_pgMar, "top", sect_dim$margins["top"])
      fill_if_missing(current_pgMar, "bottom", sect_dim$margins["bottom"])
      fill_if_missing(current_pgMar, "left", sect_dim$margins["left"])
      fill_if_missing(current_pgMar, "right", sect_dim$margins["right"])
    }
  }
  x
}

officer_section_fortify <- function(node, x) {
  title_page_tag <- "<w:titlePg xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"

  headers <- Filter(function(x) {
    xml_name(x) %in% "headerReference"
  }, xml_children(node))

  for (i in seq_along(headers)) {
    if (!inherits(xml_child(headers[[i]], "w:hdr"), "xml_missing")) {
      hof_spread_to_file(node = headers[[i]], x = x, type = "header")
    }
    if (xml_attr(headers[[i]], "type") %in% "first" && inherits(xml_child(node, "w:titlePg"), "xml_missing")) {
      xml_add_child(node, as_xml_document(title_page_tag))
    }
    if (xml_attr(headers[[i]], "type") %in% "even") {
      x$settings$even_and_odd_headers <- TRUE
    }
  }

  footers <- Filter(function(x) {
    xml_name(x) %in% "footerReference"
  }, xml2::xml_children(node))
  for (i in seq_along(footers)) {
    if (!inherits(xml_child(footers[[i]], "w:ftr"), "xml_missing")) {
      hof_spread_to_file(node = footers[[i]], x = x, type = "footer")
    }
    if (xml_attr(footers[[i]], "type") %in% "first" && inherits(xml_child(node, "w:titlePg"), "xml_missing")) {
      xml_add_child(node, as_xml_document(title_page_tag))
    }
    if (xml_attr(footers[[i]], "type") %in% "even") {
      x$settings$even_and_odd_headers <- TRUE
    }
  }
  x
}

fill_if_missing <- function(node, attr, val) {
  att_val <- xml_attr(node, attr)
  if (missing(att_val)) {
    xml_attr(node, attr) <- val
  }
  node
}


section_dimensions <- function(node) {
  section_obj <- as_list(node)

  landscape <- FALSE
  if (!is.null(attr(section_obj$pgSz, "orient")) && attr(section_obj$pgSz, "orient") == "landscape") {
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

  list(
    page = c("width" = w_ref, "height" = h_ref),
    landscape = landscape,
    margins = c(
      top = mar_t, bottom = mar_b,
      left = mar_l, right = mar_r,
      header = mar_h, footer = mar_f
    )
  )
}

hof_next_file <- function(x, type = "header") {
  pattern <- paste0("^(", type, ")([0-9]+)(\\.xml)$")
  files <- list.files(
    path = file.path(x$package_dir, "word"),
    pattern = pattern
  )
  if (length(files) < 1) {
    files <- paste0(type, "0.xml")
  }
  str_id <- gsub(pattern, "\\2", files)
  paste0(type, max(as.integer(str_id)) + 1L, ".xml")
}


hof_spread_to_file <- function(node, x, type = "header") {
  xml_basename <- hof_next_file(x, type = type)
  if ("header" %in% type) {
    selector_drop <- "w:hdr"
  } else {
    selector_drop <- "w:ftr"
  }
  node_str <- c(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    as.character(xml_child(node, 1))
  )
  writeLines(node_str, file.path(x$package_dir, "word", xml_basename),
    useBytes = TRUE
  )
  relationships <- docx_body_relationship(x)
  rid <- sprintf("rId%.0f", relationships$get_next_id())
  relationships$add(
    id = rid, type = paste0("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", type),
    target = xml_basename
  )
  xml_remove(xml_child(node, selector_drop))
  xml_attr(node, "r:id") <- rid
  x$content_type$add_override(
    setNames(
      paste0("application/vnd.openxmlformats-officedocument.wordprocessingml.", type, "+xml"),
      paste0("/word/", xml_basename)
    )
  )
}
