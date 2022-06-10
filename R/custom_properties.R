#' @importFrom xml2 xml_children
read_custom_properties <- function( package_dir ){
  filename <- file.path(package_dir, "docProps/custom.xml")
  if( !file.exists(filename) ) {
    filename <- system.file(package = "officer", "template/custom.xml")
  }
  doc <- read_xml(filename)
  all_children <- xml_children(doc)
  pid_values <- vapply(all_children, xml_attr, NA_character_, "pid")
  name_values <- vapply(all_children, xml_attr, NA_character_, "name")
  value_values <- vapply(all_children, xml_text, NA_character_)
  str <- c(pid_values, name_values, value_values)

  z <- matrix(str, ncol = 3,
    dimnames = list(NULL, c("pid", "name", "value")))
  z <- list(data = z)
  class(z) <- "custom_properties"
  z
}

`[<-.custom_properties` <- function( x, i, j, value ){
  if( !i %in% x$data[,"name"] ) {
    if(nrow(x$data)<1) {
      pid <- 2
    } else {
      pid <- max(as.integer(x$data[, "pid"])) + 1L
    }
    new <- matrix(c(as.character(pid), i, value), ncol = 3)
    x$data <- rbind(x$data, new)
  } else {
    x$data[x$data[,"name"] %in% i, j] <- value
  }
  x
}
`[.custom_properties` <- function( x, i, j ){
  if(nrow(x$data) < 1) {
    return(character())
  }
  if(missing(i)) {
    x$data[, j]
  } else {
    x$data[x$data[,"name"] %in% i, j]
  }
}

write_custom_properties <- function(custom_props, package_dir){
  xml_props <- sprintf(
    "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"%s\" name=\"%s\"><vt:lpwstr>%s</vt:lpwstr></property>",
    custom_props$data[,1],
    custom_props$data[,2],
    custom_props$data[,3])

  xml_ <- c("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
            "<Properties ",
            "xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" ",
            "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">",
            xml_props,
            "</Properties>")
  props_dir = file.path(package_dir, "docProps")
  dir.create(props_dir, recursive = TRUE, showWarnings = FALSE)
  filename <- file.path(props_dir, "custom.xml")
  writeLines(enc2utf8(xml_), filename, useBytes=TRUE)
  invisible()
}
