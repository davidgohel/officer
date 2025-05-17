xml_document_to_chrs <- function(xml_doc) {
  # Write the XML document to a temporary file
  con <- textConnection(as.character(xml_doc))
  xml_str <- readLines(con)
  close(con)
  xml_str <- gsub("^[[:blank:]]+", "", xml_str)
  xml_str
}
