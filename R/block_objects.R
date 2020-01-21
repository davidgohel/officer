#' @export
#' @title docx seqfield xml
#' @description Create a string representation of a seqfield.
#' @param sf seqfield value
#' @examples
#' seqfield(" REF label \\h ")
seqfield <- function( sf, as_scalar = TRUE ){
  xml_elt_1 <- paste0(wml_with_ns("w:r"),
                      "<w:rPr/>",
                      "<w:fldChar w:fldCharType=\"begin\"/>",
                      "</w:r>")
  xml_elt_2 <- paste0(wml_with_ns("w:r"),
                      "<w:rPr/>",
                      sprintf("<w:instrText xml:space=\"preserve\">%s</w:instrText>", sf ),
                      "</w:r>")
  xml_elt_3 <- paste0(wml_with_ns("w:r"),
                      "<w:rPr/>",
                      "<w:fldChar w:fldCharType=\"end\"/>",
                      "</w:r>")
  if( as_scalar)
    paste0(xml_elt_1, xml_elt_2, xml_elt_3)
  else
    c(xml_elt_1, xml_elt_2, xml_elt_3)
}

#' @export
#' @title docx bookmark xml
#' @description Create a string representation of a bookmark.
#' @param id bookmark name
#' @param str a scalar character representing an xml *run*.
#' @examples
#' bookmark("<r><t>hi</t></r>")
bookmark <- function(id, str){
  new_id <- UUIDgenerate()
  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\"/>", new_id, id )
  bm_start_end <- sprintf("<w:bookmarkEnd w:id=\"%s\"/>", new_id )
  paste0(bm_start_str, str, bm_start_end)
}



#' @export
#' @title numbered caption
#' @description Create a string representation of a numbered
#' caption that can be used for cross reference.
#' @param label a scalar character representing label to display
#' @param id cross reference identifier
#' @param style paragraph style identifier
numbered_caption <- function(label, style, id, seq_id = "table", pre_label = "TABLE ", post_label = ": "){

  z <- list(
    label = label,
    id = id,
    seq_id = seq_id,
    pre_label = pre_label,
    post_label = post_label
  )
  class(z) <- "numbered_caption"

  z
}

#' @export
#' @param type output format, one of docx, html.
#' @param ... unused
#' @rdname numbered_caption
format.numbered_caption = function (x, type = "docx", base_document = NULL, ...){
  stopifnot( length(type) == 1,
             type %in% c("docx", "html") )

  if( type == "docx" ){
    if( is.null(base_document)) {
      base_document <- get_reference_value("docx")
    }

    if( is.character(base_document)) {
      base_document <- read_docx(path=base_document)
    } else if( !inherits(base_document, "rdocx") ){
      stop("base_document can only be the path to a docx file or an rdocx document.")
    }

    if( is.null(x$style) )
      style <- base_document$default_styles$paragraph
    style_id <- get_style_id(data = base_document$styles, style=style, type = "paragraph")

    run_str_pre <- sprintf( "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", x$pre_label)
    run_str_post <- sprintf( "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", x$post_label)
    sf_str <- seqfield(sf = paste0("SEQ ", x$seq_id, " \u005C* Arabic \u005Cs 1 \u005C* MERGEFORMAT"))
    run_str <- sprintf( "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", htmlEscape(x$label))
    run_str <- paste0(run_str_pre, sf_str, run_str_post, run_str)
    run_str <- bookmark(x$id, run_str)
    out <- sprintf( "%s<w:pPr><w:pStyle w:val=\"%s\"/></w:pPr>%s</w:p>",
                    wml_with_ns("w:p"), style_id, run_str)
  } else stop("unimplemented")

  out
}
