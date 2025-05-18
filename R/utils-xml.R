xml_document_to_chrs <- function(xml_doc) {
  # Write the XML document to a temporary file
  con <- textConnection(as.character(xml_doc))
  xml_str <- readLines(con)
  close(con)
  xml_str <- gsub("^[[:blank:]]+", "", xml_str)
  xml_str
}

mv_ns_definitions_to_document_node <- function(xml_str) {
  m <- gregexpr(" xmlns\\:[[:alnum:]]+=[^ >]+", xml_str)
  allns <- unique(unlist(regmatches(xml_str, m)))
  regmatches(xml_str, m) <- ""
  m <- regexpr("<w\\:document ", xml_str)
  regmatches(xml_str, m) <- paste0("<w:document", paste(allns, collapse = ""), " ")
  xml_str
}

# replace body xml object with a new xml stored as a character vector
replace_xml_body_from_chr <- function(x, xml_str) {
  tf_xml <- tempfile(fileext = ".txt")
  writeLines(xml_str, tf_xml, useBytes = TRUE)
  x$doc_obj$replace_xml(tf_xml)
  x
}
