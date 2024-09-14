#' @importFrom xml2 xml_children
read_custom_properties <- function(package_dir) {
  filename <- file.path(package_dir, "docProps/custom.xml")
  if (!file.exists(filename)) {
    filename <- system.file(package = "officer", "template/custom.xml")
  }
  doc <- read_xml(filename)
  all_children <- xml_children(doc)

  pid_values <- vapply(all_children, xml_attr, NA_character_, "pid")
  name_values <- vapply(all_children, xml_attr, NA_character_, "name")
  properties_types <- vapply(all_children, function(x) {
    xml_name(xml_child(x, 1))
  }, NA_character_)

  # this complicated thing is there because some unidentifed
  # softwares fill this fields with invalid xml - so we have
  # to consider it will not be correctly rewritten and will
  # make an invalid document
  value_values <- vapply(all_children, function(x) {
    as.character(xml_child(x, 1))
  }, NA_character_)
  pat1 <- sprintf("<vt\\:(%s)/>", paste0(properties_types, collapse = "|"))
  pat2 <- sprintf("<vt\\:(%s)>", paste0(properties_types, collapse = "|"))
  pat3 <- sprintf("</vt\\:(%s)>", paste0(properties_types, collapse = "|"))
  value_values <- gsub(pat1, "", value_values)
  value_values <- gsub(pat2, "", value_values)
  value_values <- gsub(pat3, "", value_values)
  str <- c(pid_values, name_values, properties_types, value_values)
  z <- matrix(str,
    ncol = 4,
    dimnames = list(NULL, c("pid", "name", "type", "value"))
  )
  z <- list(data = z)
  class(z) <- "custom_properties"
  z
}


#' @export
`[<-.custom_properties` <- function(x, i, j, value) {
  if (!i %in% x$data[, "name"]) {
    if (nrow(x$data) < 1) {
      pid <- 2
    } else {
      pid <- max(as.integer(x$data[, "pid"])) + 1L
    }
    new <- matrix(c(as.character(pid), i, "lpwstr", value), ncol = 4)
    x$data <- rbind(x$data, new)
  } else {
    x$data[x$data[, "name"] %in% i, j] <- value
  }
  x
}


#' @export
`[.custom_properties` <- function(x, i, j) {
  if (nrow(x$data) < 1) {
    return(character())
  }
  if (missing(i)) {
    x$data[, j]
  } else {
    x$data[x$data[, "name"] %in% i, j]
  }
}

write_custom_properties <- function(custom_props, package_dir) {
  xml_props <- sprintf(
    "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"%s\" name=\"%s\"><vt:%s>%s</vt:%s></property>",
    custom_props$data[, 1],
    custom_props$data[, 2],
    custom_props$data[, 3],
    custom_props$data[, 4],
    custom_props$data[, 3]
  )

  xml_ <- c(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<Properties ",
    "xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" ",
    "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">",
    xml_props,
    "</Properties>"
  )
  props_dir <- file.path(package_dir, "docProps")
  dir.create(props_dir, recursive = TRUE, showWarnings = FALSE)
  filename <- file.path(props_dir, "custom.xml")
  writeLines(enc2utf8(xml_), filename, useBytes = TRUE)
  invisible()
}
