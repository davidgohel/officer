#' @export
#' @title add images into a docx object
#' @description reference images into a Word document.
#' This function is to be used with \code{\link{wml_link_images}}.
#'
#' Images need to be referenced into the
#' Word document, this will generate unique
#' identifiers that need to be known to
#' link these images with their corresponding xml code (wml).
#'
#' @details
#' The function copy images into the docx document and a data.frame
#' describing images and their id is returned.
#' @param x docx object
#' @param src a vector of character containing image filenames.
docx_reference_img <- function( x, src){
  src <- unique( src )

  rel_path <- file.path(x$package_dir, "word/_rels/document.xml.rels")
  relationship <- read_relationship(rel_path)

  int_id <- as.integer( gsub(pattern = "^rId", replacement = "", x = relationship$id ) )
  last_id <- as.integer( max(int_id) )

  img_path <- file.path(x$package_dir, "word", "media")
  dir.create(img_path, recursive = TRUE, showWarnings = FALSE)
  file.copy(from = src, to = file.path(x$package_dir, "word", "media", basename(src)))

  rId <- paste0("rId", seq_along(src) + last_id)

  new_rel <- data.frame(
    id = paste0("rId", seq_along(src) + last_id),
    type = rep("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
               length(src)),
    target = file.path("media", basename(src) ),
    stringsAsFactors = FALSE )

  write_relationship(x, rbind(relationship, new_rel) )
  new_rel$src <- src
  new_rel$id <- seq_along(src) + last_id
  new_rel
}

#' @importFrom xml2 as_xml_document xml_attr<-
#' @export
#' @title transform an xml string with images references
#' @description The function replace images filenames
#' in an xml string with their id. The wml code cannot
#' be valid without this operation.
#' @details
#' The function is available to let the creation of valid
#' wml code containing references to images.
#' @param str wml string
#' @param ref result of \code{\link{docx_reference_img}}
wml_link_images <- function(str, ref ){
  doc <- as_xml_document(str)
  for(id in seq_along(ref$src) ){
    xpth <- paste0("//w:drawing",
                   "[wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip",
                   sprintf( "[contains(@r:embed,'%s')]", ref$src[id]),
                   "]")
    src_nodes <- xml_find_all(doc, xpth)
    xml_find_all(src_nodes, "wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip") %>%
    {xml_attr(., "r:embed") <- paste0("rId", ref$id[id]);.}
  }
  as.character(doc)
}
