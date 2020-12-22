footnote_add_xml <- function(x, str, pos, refnote){
  pos_list <- c("after", "before", "on")
  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )


  xml_elt <- as_xml_document(str)
  ref_elt <- as_xml_document(refnote)
  cursor_elt <- x$footnotes$get_at_cursor()

  # add footnote ref as first run of first paragraph
  first_run <- xml_child(xml_child(xml_elt, 1), "w:r")
  if( !inherits(first_run, "xml_missing") ){
    xml_add_sibling(first_run, .value = ref_elt, .where = "before")
  }

  if( pos %in% "before" && x$footnotes$length() > 2 ){
    xml_add_sibling(cursor_elt, xml_elt, .where = "after")
    x$footnotes$cursor_end()
  } else if( x$footnotes$length() == 2){
    xml_add_sibling(cursor_elt, xml_elt, .where = "after")
    x$footnotes$cursor_end()
  } else if( pos %in% "on" ){
    xml_replace(cursor_elt, xml_elt)
    x$footnotes$cursor_forward()
  } else {
    xml_add_sibling(cursor_elt, xml_elt, .where = "after")
    x$footnotes$cursor_forward()
  }

  x
}

#' @export
#' @title append a footnote
#' @description append a new footnote into a paragraph of an rdocx object
#' @param x an rdocx object
#' @param style text style to be used for the reference note
#' @param blocks set of blocks to be used as footnote content returned by
#'   function \code{\link{block_list}}.
#' @param pos where to add the new element relative to the cursor, "after" or
#'   "before".
#' @examples
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' bl <- block_list(
#'   fpar(ftext("hello", shortcuts$fp_bold())),
#'   fpar(
#'     ftext("hello world", shortcuts$fp_bold()),
#'     external_img(src = img.file, height = 1.06, width = 1.39)
#'   )
#' )
#'
#' x <- read_docx()
#' x <- body_add_par(x, "Hello ", style = "Normal")
#' x <- slip_in_text(x, "world", style = "strong")
#' x <- slip_in_footnote(x, style = "reference_id", blocks = bl)
#'
#' print(x, target = tempfile(fileext = ".docx"))
slip_in_footnote <- function( x, style = NULL, blocks, pos = "after" ){

  if( !inherits(blocks, "block_list") ){
    stop("blocks must be an object created by function block_list().", call. = FALSE)
  }

  pos_list <- c("after", "before")
  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  if( is.null(style) )
    style <- x$default_styles$character

  id <- x$footnotes$length() - 1L

  style_id <- get_style_id(data = x$styles, style=style, type = "character")

  footnote_elt <- paste0( wr_ns_yes,
                          "<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>",
                          "<w:footnoteRef/></w:r>")
  footnote_elt <- sprintf(footnote_elt, style_id)

  new_src <- lapply(blocks, function(x){
    sapply(x$chunks, function(x){
      if( inherits(x, "external_img") )
        as.character(x)
      else NA_character_
    })
  })
  new_src <- unlist(new_src)
  new_src <- new_src[!is.na(new_src)]
  new_src <- unique(new_src)

  blocks <- sapply(blocks, to_wml)
  blocks <- paste(blocks, collapse = "")
  x <- part_reference_img(x, new_src, "footnotes")

  base_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""
  ftn_elt <- paste0( "<w:footnote ", base_ns,
                     sprintf(" w:id=\"%.0f\">", id),
                     blocks,
                     "</w:footnote>")
  ftn_elt <- wml_part_link_images(x, ftn_elt, "footnotes")
  x <- footnote_add_xml(x, str = ftn_elt, pos = "after", footnote_elt)

  doc_ref_elt <- paste0( wr_ns_yes,
                         "<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>",
                         "<w:footnoteReference w:id=\"%.0f\"/></w:r>")
  doc_ref_elt <- sprintf(doc_ref_elt, style_id, id)
  doc_ref_elt <- as_xml_document(doc_ref_elt)


  cursor_elt <- x$doc_obj$get_at_cursor()
  pos <- ifelse(pos=="after", length(xml_children(cursor_elt)), 1)
  xml_add_child(.x = cursor_elt, .value = doc_ref_elt, .where = pos )

  next_id <- x$doc_obj$relationship()$get_next_id()
  x$doc_obj$relationship()$add(
    paste0("rId", next_id),
    type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    target = "footnotes.xml" )
  x
}
