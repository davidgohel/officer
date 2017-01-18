#' @title read openxml relationship content
#' @description read a string or a file from a relationship file and return a detailed
#' data.frame containing imported attributes of the relationship file. These relationship
#' files are located in openxml packages.
#' @param x an xml string or a file name
#' @export
#' @importFrom xml2 read_xml xml_children xml_ns xml_attr
read_relationship <- function( x ) {
  doc <- read_xml( x = x )
  children <- xml_children( doc )
  ns <- xml_ns( doc )
  id <- sapply( children, xml_attr, attr = "Id", ns)
  type <- sapply( children, xml_attr, attr = "Type", ns)
  target <- sapply( children, xml_attr, attr = "Target", ns)
  data.frame(id = id, type = type, target = target, stringsAsFactors = FALSE )
}
