#' @export
#' @title Add comment in a 'Word' document
#' @description Add a comment at the cursor location. The comment
#' is added on the first run of text in the current paragraph.
#' @param x an rdocx object
#' @param id bookmark name
#' @examples
#'
#' # cursor_bookmark ----
#'
#' doc <- read_docx()
#' doc <- body_add_par(doc, "centered text", style = "centered")
#' doc <- body_comment(doc, )
body_comment <- function(x, value) {
  cursor_elt <- docx_current_block_xml(x)
  ns_ <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
  new_id <- uuid_generate()

  cm_start_str <- sprintf("<w:commentRangeStart w:id=\"%s\" %s/>", new_id, ns_)
  cm_start_end <- sprintf("<w:commentRangeEnd %s w:id=\"%s\"/>", ns_, new_id)

  browser()
  path_ <- paste0(xml_path(cursor_elt), "//w:r")

  node <- xml_find_first(x$doc_obj$get(), path_)
  xml_add_sibling(node, as_xml_document(cm_start_str), .where = "before")
  xml_add_sibling(node, as_xml_document(cm_start_end), .where = "after")

  x
}
