#' @export
#' @title open a connexion to a 'Word' file
#' @description read and import a docx file as an R object
#' representing the document.
#' @param path path to the docx file to use a base document.
#' @param x a docx object
#' @examples
#' # create a docx object with default template ---
#' docx()
#'
#' @importFrom xml2 read_xml xml_length xml_find_first
#' @importFrom utils unzip
docx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.docx")


  package_dir <- tempfile()
  unzip( zipfile = path, exdir = package_dir )

  document_path <- file.path(package_dir, "word/document.xml")
  doc <- read_xml(document_path)
  len <- xml_find_first(doc, "/w:document/w:body") %>% xml_length()

  obj <- list(
    package_dir = package_dir,
    cursor = sprintf("/w:document/w:body/*[%.0f]", len - 1),
    rels = relationship$new()$feed_from_xml( file.path(package_dir, "word/_rels/document.xml.rels") ),
    styles = read_styles(package_dir),
    xml_doc = doc
  )
  class(obj) <- "docx"

  obj
}

#' @export
#' @param target path to the docx file to write
#' @param ... unused
#' @rdname docx
#' @examples
#' # write a docx object in a docx file ----
#' if( require(magrittr) ){
#'   docx() %>% print(target = "out.docx")
#'   # full path of produced file is returned
#'   print(.Last.value)
#' }
#'
#' @importFrom xml2 xml_attr<- xml_find_all xml_find_all
#' @importFrom purrr walk2
print.docx <- function(x, target = NULL, ...){

  if( is.null( target) ){
    cat("docx document\n")
    print(x$styles)
    return(invisible())
  }

  if( file_ext(target) != "docx" )
    stop(target , " should have '.docx' extension.")

  document_path <- file.path(x$package_dir, "word/document.xml")

  # make all id unique
  all_uid <- xml_find_all(x$xml_doc, "//*[@id]")
  walk2(all_uid, seq_along(all_uid)-1, function(x, z) {
    xml_attr(x, "id") <- z
    x
  })

  write_xml(x$xml_doc, file = document_path)
  add_content_type(x, extension = "png", type = "image/png")
  add_content_type(x, extension = "jpeg", type = "image/jpeg")
  add_content_type(x, extension = "jpg", type = "application/octet-stream")
  add_content_type(x, extension = "emf", type = "image/x-emf")

  pack_folder(folder = x$package_dir, target = target )
}

#' @export
#' @examples
#' # how many element are there in the document ----
#' length( docx() )
#'
#' @importFrom xml2 read_xml xml_length xml_find_first
#' @rdname docx
length.docx <- function( x ){
  xml_find_first(x$xml_doc, "/w:document/w:body") %>% xml_length()

}


