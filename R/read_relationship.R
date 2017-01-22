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


# compute ids associated with images to be inserted into a MS doc
#' @export
fortify_docx_img <- function( long_str, relationships, imgs){
  imgs <- unique( imgs )

  if( length(imgs) > 0 ) {


    int_id <- as.integer( gsub(pattern = "^rId", replacement = "", x = relationships$id ) )
    last_id <- as.integer( max(int_id) )
    rids <- data.frame(rId = paste0("rId", seq_along(imgs) + last_id),
               src = imgs, nvpr_id = seq_along(imgs), doc_pr_id = seq_along(imgs),
               stringsAsFactors = FALSE )

    for(id in seq_along(rids$src) ){
      long_str <- gsub(x = long_str,
                  pattern = paste0("r:embed=\"", rids$src[id]),
                  replacement = paste0("r:embed=\"", rids$rId[id]) )
      long_str <- gsub(x = long_str, pattern = "DRAWINGOBJECTID", replacement = rids$doc_pr_id[id] )
      long_str <- gsub(x = long_str, pattern = "PICTUREID", replacement = rids$nvpr_id[id] )
    }

    extra_relationships <- data.frame(
      id = rids$rId,
      type = rep("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                 length(rids$rId)),
      target = file.path("media", basename(rids$src) ),
      stringsAsFactors = FALSE )

    attr(long_str, "relations") <- extra_relationships
    attr(long_str, "copy_files") <- rids$src
  }  else {
    attr(long_str, "relations") <- relationships[NULL]
    attr(long_str, "copy_files") <- character(0)
  }
  long_str
}

