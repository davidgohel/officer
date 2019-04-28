read_core_properties <- function( package_dir ){
  filename <- file.path(package_dir, "docProps/core.xml")
  if( !file.exists(filename) )
    stop("could not find Word document properties",
         ", please edit your document and make sure properties are existing.",
         " This can be done by filling any field in the document properties panel.", call. = FALSE)
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
  attr(props, "ns") <- unclass(ns_)
  props
}

write_core_properties <- function(core_matrix, package_dir){
  if(!is.matrix(core_matrix)){
    stop("core_properties should be stored in a character matrix.")
  }

  ns_ <- attr(core_matrix, "ns")
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
  filename <- file.path(package_dir, "docProps/core.xml")
  writeLines(enc2utf8(xml_), filename, useBytes=T)
  invisible()
}
