#' @export
#' @title add images into an rdocx object
#' @description reference images into a Word document.
#' This function is to be used with \code{\link{wml_link_images}}.
#'
#' Images need to be referenced into the
#' Word document, this will generate unique
#' identifiers that need to be known to
#' link these images with their corresponding xml code (wml).
#'
#' @param x an rdocx object
#' @param src a vector of character containing image filenames.
docx_reference_img <- function( x, src){
  src <- unique( src )
  x$doc_obj$relationship()$add_img(src, root_target = "media")

  img_path <- file.path(x$doc_obj$package_dirname(), "word", "media")
  dir.create(img_path, recursive = TRUE, showWarnings = FALSE)
  file.copy(from = src, to = file.path(x$doc_obj$package_dirname(), "word", "media", basename(src)))

  x
}

#' @importFrom xml2 as_xml_document xml_attr<-
#' @export
#' @title transform an xml string with images references
#' @description The function replace images filenames
#' in an xml string with their id. The wml code cannot
#' be valid without this operation.
#' @details
#' The function is available to allow the creation of valid
#' wml code containing references to images.
#' @param x an rdocx object
#' @param str wml string
wml_link_images <- function(x, str){
  ref <- x$doc_obj$relationship()$get_data()

  ref <- ref[ref$ext_src != "",]

  doc <- as_xml_document(str)
  for(id in seq_along(ref$ext_src) ){

    xpth <- paste0("//w:drawing",
                   "[wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip",
                   sprintf( "[contains(@r:embed,'%s')]", ref$ext_src[id]),
                   "]")

    src_nodes <- xml_find_all(doc, xpth)
    xml_find_all(src_nodes, "wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip") %>%
    {xml_attr(., "r:embed") <- ref$id[id];.}
  }
  as.character(doc)
}
