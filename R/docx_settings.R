#' @export
#' @title Set 'Microsoft Word' Document Settings
#' @description
#' Set various settings of a 'Microsoft Word' document generated with 'officer'.
#' Options include:
#' - zoom factor (default view in Word),
#' - default tab stop,
#' - hyphenation zone,
#' - decimal symbol,
#' - list separator (see details below),
#' - compatibility mode,
#' - even and odd headers management (see details below),
#' - and auto hyphenation activation.
#'
#' @details
#' - `even_and_odd_headers`: If TRUE, 'Microsoft Word' will use different
#' headers for odd and even pages ("Different Odd & Even Pages" feature in Word).
#' This is useful for professional documents or reports that require alternating
#' page layouts.
#'
#' - `list_separator`: Sets the character used by 'Microsoft Word' to separate
#' items in lists (for example, when inserting tables or lists in Word). This
#' parameter affects how 'Microsoft Word' handles data import/export (CSV, etc.) and can be
#' adapted to language or local conventions (e.g., ";" for French, "," for English).
#'
#' @seealso [read_docx()]
#' @param x an rdocx object
#' @param zoom zoom factor, default is 1 (100%)
#' @param default_tab_stop default tab stop in inches, default is 0.5
#' @param hyphenation_zone hyphenation zone in inches, default is 0.25
#' @param decimal_symbol decimal symbol, default is "."
#' @param list_separator list separator, default is ";". Sets the separator
#' used by Word for lists (see details).
#' @param compatibility_mode compatibility mode, default is "15"
#' @param even_and_odd_headers whether to use different headers for even and
#' odd pages, default is FALSE. Enables the "Different Odd and Even Pages"
#' feature in 'Microsoft Word'.
#' @param auto_hyphenation whether to enable auto hyphenation, default is FALSE.
#' @param unit unit for `default_tab_stop` and `hyphenation_zone`, one of "in", "cm", "mm".
#' @examples
#' txt_lorem <- rep(
#'   "Purus lectus eros metus turpis mattis platea praesent sed. ",
#'   50
#' )
#' txt_lorem <- paste0(txt_lorem, collapse = "")
#'
#' header_first <- block_list(fpar(ftext("text for first page header")))
#' header_even <- block_list(fpar(ftext("text for even page header")))
#' header_default <- block_list(fpar(ftext("text for default page header")))
#' footer_first <- block_list(fpar(ftext("text for first page footer")))
#' footer_even <- block_list(fpar(ftext("text for even page footer")))
#' footer_default <- block_list(fpar(ftext("text for default page footer")))
#'
#' ps <- prop_section(
#'   header_default = header_default, footer_default = footer_default,
#'   header_first = header_first, footer_first = footer_first,
#'   header_even = header_even, footer_even = footer_even
#' )
#'
#' x <- read_docx()
#'
#' x <- docx_set_settings(
#'   x = x,
#'   zoom = 2,
#'   list_separator = ",",
#'   even_and_odd_headers = TRUE
#' )
#'
#' for (i in 1:20) {
#'   x <- body_add_par(x, value = txt_lorem)
#' }
#' x <- body_set_default_section(
#'   x,
#'   value = ps
#' )
#' print(x, target = tempfile(fileext = ".docx"))
docx_set_settings <- function(
  x,
  zoom = 1,
  default_tab_stop = .5,
  hyphenation_zone = .25,
  decimal_symbol = ".",
  list_separator = ";",
  compatibility_mode = "15",
  even_and_odd_headers = FALSE,
  auto_hyphenation = FALSE,
  unit = "in"
) {
  settings <- docx_settings(
    zoom = zoom,
    default_tab_stop = convin(unit = unit, x = default_tab_stop),
    hyphenation_zone = convin(unit = unit, x = hyphenation_zone),
    decimal_symbol = decimal_symbol,
    list_separator = list_separator,
    even_and_odd_headers = even_and_odd_headers,
    auto_hyphenation = auto_hyphenation,
    compatibility_mode = compatibility_mode
  )
  x$settings <- settings
  x
}

docx_settings <- function(
  zoom = 1,
  default_tab_stop = .5,
  hyphenation_zone = .25,
  decimal_symbol = ".",
  list_separator = ";",
  compatibility_mode = "15",
  even_and_odd_headers = FALSE,
  auto_hyphenation = FALSE
) {
  x <- list(
    zoom = zoom,
    default_tab_stop = default_tab_stop,
    hyphenation_zone = hyphenation_zone,
    decimal_symbol = decimal_symbol,
    list_separator = list_separator,
    even_and_odd_headers = even_and_odd_headers,
    auto_hyphenation = auto_hyphenation,
    compatibility_mode = compatibility_mode
  )
  class(x) <- "docx_settings"
  x
}

update_docx_settings_from_file <- function(x, file) {
  if (is.null(file)) {
    cli::cli_abort(
      "File settings for Word document is null."
    )
  }
  if (!file.exists(file)) {
    cli::cli_abort(
      "File settings for Word document does not exists {.file {file}}."
    )
  }

  node_doc <- read_xml(file)

  node_zoom <- xml_child(node_doc, "w:zoom")
  if (!inherits(node_zoom, "xml_missing")) {
    x$zoom <- as.integer(xml_attr(node_zoom, "percent")) / 100
  }

  node_tab_stop <- xml_child(node_doc, "w:defaultTabStop")
  if (!inherits(node_tab_stop, "xml_missing")) {
    x$default_tab_stop <- as.integer(xml_attr(node_tab_stop, "val")) /
      1440
  }

  node_hyphenation_zone <- xml_child(node_doc, "w:hyphenationZone")
  if (!inherits(node_hyphenation_zone, "xml_missing")) {
    x$hyphenation_zone <- as.integer(xml_attr(
      node_hyphenation_zone,
      "val"
    )) /
      1440
  }

  node_auto_hyphenation <- xml_child(node_doc, "w:autoHyphenation")
  if (!inherits(node_auto_hyphenation, "xml_missing")) {
    x$auto_hyphenation <- TRUE
  }

  node_decimal_symbol <- xml_child(node_doc, "w:decimalSymbol")
  if (!inherits(node_decimal_symbol, "xml_missing")) {
    x$decimal_symbol <- xml_attr(node_decimal_symbol, "val")
  }

  node_list_separator <- xml_child(node_doc, "w:listSeparator")
  if (!inherits(node_list_separator, "xml_missing")) {
    x$list_separator <- xml_attr(node_list_separator, "val")
  }

  node_evenodd_headers <- xml_child(node_doc, "w:evenAndOddHeaders")
  if (!inherits(node_evenodd_headers, "xml_missing")) {
    x$even_and_odd_headers <- TRUE
  } else {
    x$even_and_odd_headers <- FALSE
  }

  x
}

write_docx_settings <- function(x) {
  settings <- x$settings
  str <- paste0(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
    sprintf("<w:zoom w:percent=\"%.0f\"/>", settings$zoom * 100),
    sprintf("<w:defaultTabStop w:val=\"%.0f\"/>", settings$default_tab_stop * 1440),
    if (settings$auto_hyphenation) sprintf("<w:autoHyphenation/>"),
    sprintf("<w:hyphenationZone w:val=\"%.0f\"/>", settings$hyphenation_zone * 1440),
    sprintf(
      "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"%s\"/></w:compat>",
      settings$compatibility_mode
    ),
    sprintf("<w:decimalSymbol w:val=\"%s\"/>", settings$decimal_symbol),
    sprintf("<w:listSeparator w:val=\"%s\"/>", settings$list_separator),
    if (settings$even_and_odd_headers) {
      "<w:evenAndOddHeaders w:val=\"1\"/>"
    } else {
      "<w:evenAndOddHeaders w:val=\"0\"/>"
    },
    "</w:settings>"
  )
  file <- file.path(x$package_dir, "word", "settings.xml")
  writeLines(str, file, useBytes = TRUE)
  TRUE
}
