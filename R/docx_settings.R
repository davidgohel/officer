docx_settings <- function(zoom = 1,
                          default_tab_stop = .5,
                          hyphenation_zone = .25,
                          decimal_symbol = ".",
                          list_separator = ";",
                          compatibility_mode = "15",
                          even_and_odd_headers = FALSE
                          ) {
  x <- list(
    zoom = zoom,
    default_tab_stop = default_tab_stop,
    hyphenation_zone = hyphenation_zone,
    decimal_symbol = decimal_symbol,
    list_separator = list_separator,
    even_and_odd_headers = even_and_odd_headers,
    compatibility_mode = compatibility_mode
  )
  class(x) <- "docx_settings"
  x
}

#' @export
update.docx_settings <- function(object,
                                 zoom = NULL,
                                 default_tab_stop = NULL,
                                 hyphenation_zone = NULL,
                                 decimal_symbol = NULL,
                                 list_separator = NULL,
                                 compatibility_mode = NULL,
                                 even_and_odd_headers = NULL,
                                 file = NULL,
                                 ...) {
  if (!is.null(zoom)) {
    object$zoom <- zoom
  }
  if (!is.null(default_tab_stop)) {
    object$default_tab_stop <- default_tab_stop
  }
  if (!is.null(hyphenation_zone)) {
    object$hyphenation_zone <- hyphenation_zone
  }
  if (!is.null(decimal_symbol)) {
    object$decimal_symbol <- decimal_symbol
  }
  if (!is.null(list_separator)) {
    object$list_separator <- list_separator
  }
  if (!is.null(compatibility_mode)) {
    object$compatibility_mode <- compatibility_mode
  }
  if (!is.null(even_and_odd_headers)) {
    object$even_and_odd_headers <- even_and_odd_headers
  }
  if (!is.null(file) && file.exists(file)) {
    node_doc <- read_xml(file)

    node_zoom <- xml_child(node_doc, "w:zoom")
    if (!inherits(node_zoom, "xml_missing")) {
      object$zoom <- as.integer(xml_attr(node_zoom, "percent")) / 100
    }

    node_tab_stop <- xml_child(node_doc, "w:defaultTabStop")
    if (!inherits(node_tab_stop, "xml_missing")) {
      object$default_tab_stop <- as.integer(xml_attr(node_tab_stop, "val")) / 1440
    }

    node_hyphenation_zone <- xml_child(node_doc, "w:hyphenationZone")
    if (!inherits(node_hyphenation_zone, "xml_missing")) {
      object$hyphenation_zone <- as.integer(xml_attr(node_hyphenation_zone, "val")) / 1440
    }

    node_decimal_symbol <- xml_child(node_doc, "w:decimalSymbol")
    if (!inherits(node_decimal_symbol, "xml_missing")) {
      object$decimal_symbol <- xml_attr(node_decimal_symbol, "val")
    }

    node_list_separator <- xml_child(node_doc, "w:listSeparator")
    if (!inherits(node_list_separator, "xml_missing")) {
      object$list_separator <- xml_attr(node_list_separator, "val")
    }
  }
  object
}

#' @export
to_wml.docx_settings <- function(x, add_ns = FALSE, ...) {
  out <- paste0(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
    sprintf("<w:zoom w:percent=\"%.0f\"/>", x$zoom * 100),
    sprintf("<w:defaultTabStop w:val=\"%.0f\"/>", x$default_tab_stop * 1440),
    sprintf("<w:hyphenationZone w:val=\"%.0f\"/>", x$hyphenation_zone * 1440),
    sprintf("<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"%s\"/></w:compat>", x$compatibility_mode),
    sprintf("<w:decimalSymbol w:val=\"%s\"/>", x$decimal_symbol),
    sprintf("<w:listSeparator w:val=\"%s\"/>", x$list_separator),
    if (x$even_and_odd_headers) "<w:evenAndOddHeaders/>",
    "</w:settings>"
  )
  out
}

write_docx_settings <- function(x) {
  str <- to_wml(x$settings)
  file <- file.path(x$package_dir, "word", "settings.xml")
  writeLines(str, file, useBytes = TRUE)
  TRUE
}
