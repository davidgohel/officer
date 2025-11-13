#' @export
#' @title Read 'Word' styles
#' @description Read Word styles and get results in
#' a data.frame.
#' @param x an rdocx object
#' @param type,is_default subsets for types (i.e. paragraph) and
#' default style (when `is_default` is TRUE or FALSE)
#' @examples
#' x <- read_docx()
#' styles_info(x)
#' styles_info(x, type = "paragraph", is_default = TRUE)
#' @family functions for Word document informations
styles_info <- function(
  x,
  type = c("paragraph", "character", "table", "numbering"),
  is_default = c(TRUE, FALSE)
) {
  styles <- x$styles
  styles <- styles[
    styles$style_type %in% type & styles$is_default %in% is_default,
  ]
  styles
}


#' @export
#' @title Add or replace paragraph style in a Word document
#' @description The function lets you add or replace a Word paragraph style.
#' @param x an rdocx object
#' @param style_id a unique style identifier for Word.
#' @param style_name a unique label associated with the style identifier.
#' This label is the name of the style when Word edit the document.
#' @param base_on the style name used as base style
#' @param fp_p paragraph formatting properties, see [fp_par()].
#' @param fp_t default text formatting properties. This is used as
#' text formatting properties, see [fp_text()]. If NULL (default), the
#' paragraph will used the default text formatting properties (defined by
#' the `base_on` argument).
#' @examples
#' library(officer)
#'
#' doc <- read_docx()
#'
#' doc <- docx_set_paragraph_style(
#'   doc,
#'   style_id = "rightaligned",
#'   style_name = "Explicit label",
#'   fp_p = fp_par(text.align = "right", padding = 20),
#'   fp_t = fp_text_lite(
#'     bold = TRUE,
#'     shading.color = "#FD34F0",
#'     color = "white")
#' )
#'
#' doc <- body_add_par(doc,
#'   value = "This is a test",
#'   style = "Explicit label")
#'
#' docx_file <- print(doc, target = tempfile(fileext = ".docx"))
#' docx_file
docx_set_paragraph_style <- function(
  x,
  style_id,
  style_name,
  base_on = "Normal",
  fp_p = fp_par(),
  fp_t = NULL
) {
  styles_file <- file.path(x$package_dir, "word/styles.xml")
  doc <- read_xml(styles_file)

  if (grepl("[^a-zA-Z0-9\\-]+", style_id)) {
    stop("`style_id` should only contain '-', numbers and ascii characters.")
  }

  node_styles <- xml_find_first(doc, "/w:styles")

  fp_p$word_style <- NULL

  if (!is.null(fp_t)) {
    fp_t_xml <- rpr_wml(fp_t)
  } else {
    fp_t_xml <- ""
  }
  base_on <- get_style_id(data = x$styles, style = base_on, type = "paragraph")

  xml_code <- paste0(
    sprintf(
      "<w:style xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"%s\">",
      style_id
    ),
    sprintf("<w:name w:val=\"%s\"/>", style_name),
    sprintf("<w:basedOn w:val=\"%s\"/>", base_on),
    ppr_wml(fp_p),
    fp_t_xml,
    "</w:style>"
  )

  node_style <- xml_child(
    node_styles,
    sprintf("w:style[@w:styleId='%s']", style_id)
  )
  if (inherits(node_style, "xml_missing")) {
    xml_add_child(node_styles, as_xml_document(xml_code))
  } else {
    xml_replace(node_style, as_xml_document(xml_code))
  }

  write_xml(doc, file = styles_file)
  styles <- read_docx_styles(x$package_dir)
  x$styles <- styles

  x
}

#' @export
#' @title Add character style in a Word document
#' @description The function lets you add or modify Word character styles.
#' @param x an rdocx object
#' @param style_id a unique style identifier for Word.
#' @param style_name a unique label associated with the style identifier.
#' This label is the name of the style when Word edit the document.
#' @param base_on the character style name used as base style
#' @param fp_t Text formatting properties, see [fp_text()].
#' @examples
#' library(officer)
#' doc <- read_docx()
#'
#' doc <- docx_set_character_style(
#'   doc,
#'   style_id = "newcharstyle",
#'   style_name = "label for char style",
#'   base_on = "Default Paragraph Font",
#'   fp_text_lite(
#'     shading.color = "red",
#'     color = "white")
#' )
#' paragraph <- fpar(
#'   run_wordtext("hello",
#'     style_id = "newcharstyle"))
#'
#' doc <- body_add_fpar(doc, value = paragraph)
#' docx_file <- print(doc, target = tempfile(fileext = ".docx"))
#' docx_file
docx_set_character_style <- function(
  x,
  style_id,
  style_name,
  base_on,
  fp_t = fp_text_lite()
) {
  styles_file <- file.path(x$package_dir, "word/styles.xml")
  doc <- read_xml(styles_file)
  node_styles <- xml_find_first(doc, "/w:styles")

  if (grepl("[^a-zA-Z0-9\\-]+", style_id)) {
    stop("`style_id` should only contain '-', numbers and ascii characters.")
  }

  base_on <- get_style_id(data = x$styles, style = base_on, type = "character")

  xml_code <- paste0(
    sprintf(
      "<w:style xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:type=\"character\" w:customStyle=\"1\" w:styleId=\"%s\">",
      style_id
    ),
    sprintf("<w:name w:val=\"%s\"/>", style_name),
    sprintf("<w:basedOn w:val=\"%s\"/>", base_on),
    rpr_wml(fp_t),
    "</w:style>"
  )

  node_style <- xml_child(
    node_styles,
    sprintf("w:style[@w:styleId='%s']", style_id)
  )
  if (inherits(node_style, "xml_missing")) {
    xml_add_child(node_styles, as_xml_document(xml_code))
  } else {
    xml_replace(node_style, as_xml_document(xml_code))
  }

  write_xml(doc, file = styles_file)
  styles <- read_docx_styles(x$package_dir)
  x$styles <- styles

  x
}


#' @export
#' @title Replace styles in a 'Word' Document
#' @description Replace styles with others in a 'Word' document. This function
#' can be used for paragraph, run/character and table styles.
#' @param x an rdocx object
#' @param mapstyles a named list, names are the replacement style,
#' content (as a character vector) are the styles to be replaced.
#' Use [styles_info()] to display available styles.
#' @examples
#' # creating a sample docx so that we can illustrate how
#' # to change styles
#' doc_1 <- read_docx()
#'
#' doc_1 <- body_add_par(doc_1, "A title", style = "heading 1")
#' doc_1 <- body_add_par(doc_1, "Another title", style = "heading 2")
#' doc_1 <- body_add_par(doc_1, "Hello world!", style = "Normal")
#' file <- print(doc_1, target = tempfile(fileext = ".docx"))
#'
#' # now we can illustrate how
#' # to change styles with `change_styles`
#' doc_2 <- read_docx(path = file)
#' mapstyles <- list(
#'   "centered" = c("Normal", "heading 2"),
#'   "strong" = "Default Paragraph Font"
#' )
#' doc_2 <- change_styles(doc_2, mapstyles = mapstyles)
#' print(doc_2, target = tempfile(fileext = ".docx"))
change_styles <- function(x, mapstyles) {
  if (is.null(mapstyles) || length(mapstyles) < 1) {
    return(x)
  }

  table_styles <- styles_info(x, type = c("paragraph", "character", "table"))

  from_styles <- unique(as.character(unlist(mapstyles)))
  to_styles <- unique(names(mapstyles))

  if (any(is.na(mfrom <- match(from_styles, table_styles$style_name)))) {
    stop(
      "could not find style ",
      paste0(shQuote(from_styles[is.na(mfrom)]), collapse = ", "),
      ".",
      call. = FALSE
    )
  }
  if (any(is.na(mto <- match(to_styles, table_styles$style_name)))) {
    stop(
      "could not find style ",
      paste0(shQuote(to_styles[is.na(mto)]), collapse = ", "),
      ".",
      call. = FALSE
    )
  }

  mapping <- mapply(
    function(from, to) {
      id_to <- which(table_styles$style_name %in% to)
      id_to <- table_styles$style_id[id_to]

      id_from <- which(table_styles$style_name %in% from)
      types <- substring(table_styles$style_type[id_from], first = 1, last = 1)
      types[types %in% "c"] <- "r"
      types[types %in% "t"] <- "tbl"
      id_from <- table_styles$style_id[id_from]

      data.frame(
        from = id_from,
        to = rep(id_to, length(from)),
        types = types,
        stringsAsFactors = FALSE
      )
    },
    mapstyles,
    names(mapstyles),
    SIMPLIFY = FALSE
  )

  mapping <- do.call(rbind, mapping)
  row.names(mapping) <- NULL

  for (i in seq_len(nrow(mapping))) {
    all_nodes <- xml_find_all(
      x$doc_obj$get(),
      sprintf("//w:%sStyle[@w:val='%s']", mapping$types[i], mapping$from[i])
    )
    xml_attr(all_nodes, "w:val") <- rep(mapping$to[i], length(all_nodes))
  }

  x
}
