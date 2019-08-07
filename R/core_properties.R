read_core_properties <- function( package_dir ){
  filename <- file.path(package_dir, "docProps/core.xml")
  if( !file.exists(filename) ) {
    # docx and pptx files exported from Google Docs or Slides do not have properties set. Use one from the template to get by
    warning("No properties found. Using properties from template file in officer.")
    filename <- system.file(package = "officer", "template/core.xml")
  }
  doc <- read_xml(filename)
  ns_ <- xml_ns(doc)

  all_ <- xml_find_all(doc, "/cp:coreProperties/*")

  names_ <- sapply(all_, xml_name, ns = ns_)
  names_ <- strsplit(names_, ":")

  # concat all attrs in single chars
  attrs <- sapply( xml_attrs(all_), function(ns_) {
    paste0('xsi:', names(ns_), '=\"', ns_, '\"', collapse = " ")
  })
  attrs <- ifelse(sapply( xml_attrs(all_), length )<1, "", paste0(" ", attrs))

  propnames <- sapply(names_, function(x) x[2] )

  props <- matrix(c(sapply(names_, function(x) x[1] ), propnames, attrs, xml_text(all_)), ncol = 4 )
  colnames(props) <- c("ns", "name", "attrs", "value")
  rownames(props) <- propnames
  z <- list(data = props, ns = unclass(ns_))
  class(z) <- "core_properties"
  z
}

default_cp_tags <- list(
  title = list(ns = "dc", attrs = NA_character_),
  subject = list(ns = "dc", attrs = NA_character_),
  creator = list(ns = "dc", attrs = NA_character_),
  keywords = list(ns = "cp", attrs = NA_character_),
  description = list(ns = "dc", attrs = NA_character_),
  lastModifiedBy = list(ns = "cp", attrs = NA_character_),
  revision = list(ns = "cp", attrs = NA_character_),
  created = list(ns = "dcterms", attrs = 'xsi:type="dcterms:W3CDTF"'),
  modified = list(ns = "dcterms", attrs = 'xsi:type="dcterms:W3CDTF"'),
  category = list(ns = "cp", attrs = NA_character_)
)

`[<-.core_properties` <- function( x, i, j, value ){
  if( !i %in% x$data[,"name"] ) {
    attrs <- ifelse(is.na(default_cp_tags[[i]]$attrs), "", paste0(" ", default_cp_tags[[i]]$attrs))
    new <- matrix( c(default_cp_tags[[i]]$ns, i, attrs, value), nrow = 1, dimnames = list(i, colnames(x$data)))
    x$data <- rbind(x$data, new)
  } else {
    x$data[i,j] <- value
  }
  x
}
`[.core_properties` <- function( x, i, j ){
  x$data[i,j]
}

write_core_properties <- function(core_matrix, package_dir){
  ns_ <- core_matrix$ns
  core_matrix <- core_matrix$data
  if(!is.matrix(core_matrix)){
    stop("core_properties should be stored in a character matrix.")
  }

  ns_ <- paste0('xmlns:', names(ns_), '=\"', ns_, '\"', collapse = " ")
  xml_ <- paste0('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '\n',
                 '<cp:coreProperties ', ns_, '>')
  properties <- sprintf("<%s:%s%s>%s</%s:%s>",
                        core_matrix[, "ns"], core_matrix[, "name"],
                        core_matrix[, "attrs"],
                        core_matrix[, "value"],
                        core_matrix[, "ns"], core_matrix[, "name"]
  )
  xml_ <- paste0(xml_, paste0(properties, collapse = ""), "</cp:coreProperties>" )
  props_dir = file.path(package_dir, "docProps")
  dir.create(props_dir, recursive = TRUE, showWarnings = FALSE)
  filename <- file.path(props_dir, "core.xml")
  writeLines(enc2utf8(xml_), filename, useBytes=TRUE)
  invisible()
}
