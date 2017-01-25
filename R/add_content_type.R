#' @export
#' @title declare a new content type
#' @description The function add a new content type associated to an
#' extension into a docx object.
#' @param x a docx object
#' @param extension extension string value (i.e. 'png')
#' @param type corresponding type (i.e. 'image/png')
#' @examples
#' x <- docx()
#' x <- add_content_type(x, extension = "png", type = "image/png")
#' @importFrom xml2 read_xml xml_find_all xml_attr xml_add_child write_xml
add_content_type <- function( x, extension, type ){
  content_type_filename <- file.path(x$package_dir, "[Content_Types].xml")
  doc <- read_xml(x = content_type_filename )
  all_ext <- xml_find_all(doc, "//*[@Extension]")

  if( !extension %in% xml_attr(all_ext, "Extension") ){
    str <- sprintf("<Default Extension=\"%s\" ContentType=\"%s\"/>", extension, type)
    xml_add_child(doc, as_xml_document(str) )
  }

  write_xml(doc, content_type_filename )
  x
}

