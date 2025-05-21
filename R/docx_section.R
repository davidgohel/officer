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
body_end_section_landscape <- function(x, w = 16838 / 1440, h = 11906 / 1440) {
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
body_end_section_portrait <- function(x, w = 16838 / 1440, h = 11906 / 1440) {
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
body_end_section_columns <- function(
  x,
  widths = c(2.5, 2.5),
  space = .25,
  sep = FALSE
) {
  bs <- block_section(prop_section(
    section_columns = section_columns(
      widths = widths,
      space = space,
      sep = sep
    ),
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
body_end_section_columns_landscape <- function(
  x,
  widths = c(2.5, 2.5),
  space = .25,
  sep = FALSE,
  w = 16838 / 1440,
  h = 11906 / 1440
) {
  bs <- block_section(prop_section(
    page_size = page_size(width = w, height = h, orient = "landscape"),
    section_columns = section_columns(
      widths = widths,
      space = space,
      sep = sep
    ),
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
process_sections_content <- function(x, xml_str) {
  hof_summary <- extract_hof(x, xml_str)

  relationships <- docx_body_relationship(x)

  for (i in seq_len(nrow(hof_summary))) {
    # write file
    xml_doc <- as_xml_document(hof_summary[[i, "content_section"]])
    write_xml(
      xml_doc,
      file = file.path(x$package_dir, "word", hof_summary[[i, "file_name"]])
    )

    relationships$add(
      id = hof_summary[[i, "rid"]],
      type = paste0(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
        hof_summary[[i, "type"]]
      ),
      target = hof_summary[[i, "file_name"]]
    )
    x$content_type$add_override(
      setNames(
        paste0(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.",
          hof_summary[[i, "type"]],
          "+xml"
        ),
        paste0("/word/", hof_summary[[i, "file_name"]])
      )
    )
  }

  xml_str[hof_summary$reference_position] <- sprintf(
    xml_str[hof_summary$reference_position],
    hof_summary[["rid"]]
  )
  # xml_str <- gsub(" w:officer=\"true\"",  "", xml_str)

  # drop added sections contents from document xml
  lines_index_to_drop <- mapply(
    function(start, stop, str) {
      start:stop
    },
    start = hof_summary$start,
    stop = hof_summary$stop,
    SIMPLIFY = FALSE
  )
  lines_index_to_drop <- unlist(lines_index_to_drop)
  if (length(lines_index_to_drop) > 0) {
    xml_str <- xml_str[-lines_index_to_drop]
  }

  xml_str
}

guess_and_set_even_and_odd_headers <- function(x, xml_str) {
  test_even <- any(grepl(
    "(headerReference|footerReference)([^>]+)(w:type=\"even\")",
    xml_str
  ))
  if (test_even) {
    x$settings$even_and_odd_headers <- TRUE
  } else {
    x$settings$even_and_odd_headers <- FALSE
  }
  x
}

section_dimensions <- function(node) {
  section_obj <- as_list(node)

  landscape <- FALSE
  if (
    !is.null(attr(section_obj$pgSz, "orient")) &&
      attr(section_obj$pgSz, "orient") == "landscape"
  ) {
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
      top = mar_t,
      bottom = mar_b,
      left = mar_l,
      right = mar_r,
      header = mar_h,
      footer = mar_f
    )
  )
}

hof_next_file_index <- function(x, type = "header") {
  pattern <- paste0("^(", type, ")([0-9]+)(\\.xml)$")
  files <- list.files(
    path = file.path(x$package_dir, "word"),
    pattern = pattern
  )
  empty_index <- 0L
  m <- regexpr("[0-9]+", files)
  index_chars <- regmatches(files, m)
  index_ints <- as.integer(index_chars)
  index <- max(c(empty_index, index_ints), na.rm = TRUE)
  index + 1L
}


extract_hof <- function(x, xml_str) {
  # init next_rel_id that will be the first rid
  relationships <- docx_body_relationship(x)
  next_rel_id <- relationships$get_next_id()

  next_footer_file_index <- hof_next_file_index(x, type = "footer")
  next_header_file_index <- hof_next_file_index(x, type = "header")

  # footers extraction
  starts <- grep("<w:ftr", xml_str, fixed = TRUE)
  stops <- grep("</w:ftr>", xml_str, fixed = TRUE)

  content_section_chrs <- mapply(
    function(start, stop, str) {
      z <- str[start:stop]
      paste(z, collapse = "")
    },
    start = starts,
    stop = stops,
    MoreArgs = list(str = xml_str),
    SIMPLIFY = FALSE
  )
  content_section_chrs <- unlist(content_section_chrs)
  file_indexes <- seq(
    from = next_footer_file_index,
    along.with = content_section_chrs
  )
  file_names <- sprintf("%s%.0f.xml", "footer", file_indexes)

  rids <- sprintf(
    "rId%.0f",
    seq(from = next_rel_id, along.with = content_section_chrs)
  )

  reference_positions <- grep(
    "<w:footerReference w:officer",
    xml_str,
    fixed = TRUE
  )

  footer_sections <- data.frame(
    content_section = content_section_chrs,
    file_name = file_names,
    rid = rids,
    type = rep("footer", length(rids)),
    start = starts,
    stop = stops,
    reference_position = reference_positions,
    stringsAsFactors = FALSE
  )

  # headers extraction
  starts <- grep("<w:hdr", xml_str, fixed = TRUE)
  stops <- grep("</w:hdr>", xml_str, fixed = TRUE)

  content_section_chrs <- mapply(
    function(start, stop, str) {
      z <- str[start:stop]
      paste(z, collapse = "")
    },
    start = starts,
    stop = stops,
    MoreArgs = list(str = xml_str),
    SIMPLIFY = FALSE
  )
  content_section_chrs <- unlist(content_section_chrs)
  file_indexes <- seq(
    from = next_header_file_index,
    along.with = content_section_chrs
  )
  file_names <- sprintf("%s%.0f.xml", "header", file_indexes)

  rids <- sprintf(
    "rId%.0f",
    seq(
      from = next_rel_id + nrow(footer_sections),
      along.with = content_section_chrs
    )
  )

  reference_positions <- grep(
    "<w:headerReference w:officer",
    xml_str,
    fixed = TRUE
  )

  header_sections <- data.frame(
    content_section = content_section_chrs,
    file_name = file_names,
    rid = rids,
    type = rep("header", length(rids)),
    start = starts,
    stop = stops,
    reference_position = reference_positions,
    stringsAsFactors = FALSE
  )

  rbind(header_sections, footer_sections)
}



copy_header_references_everywhere <- function(x, xml_str) {
  xml_body <- docx_body_xml(x)

  all_ref <- xml_find_all(xml_body, "/w:document/w:body/w:sectPr/w:headerReference")
  all_ref <- as.character(all_ref)
  all_ref <- paste0(all_ref, collapse = "")
  m <- regexpr("(?=</w:sectPr)", xml_str, perl = TRUE)

  default_sect_pos <- tail(which(m>-1), n = 1)
  attr(m,"match.length")[default_sect_pos] <- -1
  m[default_sect_pos] <- -1

  regmatches(xml_str, m) <- all_ref

  xml_str
}

copy_footer_references_everywhere <- function(x, xml_str) {
  xml_body <- docx_body_xml(x)

  all_ref <- xml_find_all(xml_body, "/w:document/w:body/w:sectPr/w:footerReference")
  all_ref <- as.character(all_ref)
  all_ref <- paste0(all_ref, collapse = "")
  m <- regexpr("(?=</w:sectPr)", xml_str, perl = TRUE)

  default_sect_pos <- tail(which(m>-1), n = 1)
  attr(m,"match.length")[default_sect_pos] <- -1
  m[default_sect_pos] <- -1

  regmatches(xml_str, m) <- all_ref

  xml_str
}
