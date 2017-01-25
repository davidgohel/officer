#' @export
#' @title append seq field
#' @description append seq field into a paragraph of a docx object
#' @param x a docx device
#' @param str seq field value
#' @param style text style
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @examples
#' library(magrittr)
#' docx() %>%
#'   docx_add_par("Time is: ", style = "Normal") %>%
#'   docx_append_seqfield(
#'     str = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT",
#'     style = 'strong') %>%
#'
#'   docx_add_par(" - This is a figure title", style = "centered") %>%
#'   docx_append_seqfield(str = "SEQ Figure \u005C* roman",
#'     style = 'Default Paragraph Font', pos = "before") %>%
#'   docx_append_run("Figure: ", style = "strong", pos = "before") %>%
#'
#'   docx_add_par(" - This is another figure title", style = "centered") %>%
#'   docx_append_seqfield(str = "SEQ Figure \u005C* roman",
#'     style = 'strong', pos = "before")  %>%
#'   docx_append_run("Figure: ", style = "strong", pos = "before") %>%
#'   docx_add_par("This is a symbol: ", style = "Normal") %>%
#'   docx_append_seqfield(str = "SYMBOL 100 \u005Cf Wingdings",
#'     style = 'strong') %>%
#'   print(target = "seqfield.docx")
docx_append_seqfield <- function( x, str, style = "Normal", pos = "after" ){

  style_id <- get_style_id(x=x, style=style, type = "character")

  xml_elt_1 <- paste0(sprintf("<w:r %s>", base_ns),
                      sprintf("<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>", style_id),
                      "<w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/>",
                      "</w:r>")
  xml_elt_2 <- paste0(sprintf("<w:r %s>", base_ns),
                      sprintf("<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>", style_id),
                      sprintf("<w:instrText xml:space=\"preserve\" w:dirty=\"true\">%s</w:instrText>", str ),
                      "</w:r>")
  xml_elt_3 <- paste0(sprintf("<w:r %s>", base_ns),
                      sprintf("<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>", style_id),
                      "<w:r><w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/></w:r>",
                      "</w:r>")
  if( pos == "after"){
    add_xml_run(x = x, str = xml_elt_1, pos = pos)
    add_xml_run(x = x, str = xml_elt_2, pos = pos)
    add_xml_run(x = x, str = xml_elt_3, pos = pos)
  } else {
    add_xml_run(x = x, str = xml_elt_3, pos = pos)
    add_xml_run(x = x, str = xml_elt_2, pos = pos)
    add_xml_run(x = x, str = xml_elt_1, pos = pos)
  }

}

#' @export
#' @title append text
#' @description append text into a paragraph of a docx object
#' @param x a docx device
#' @param str text
#' @param style text style
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @examples
#' library(magrittr)
#' docx() %>%
#'   docx_add_par("Hello ", style = "Normal") %>%
#'   docx_append_run("world", style = "strong") %>%
#'   docx_append_run("Message is", style = "strong", pos = "before") %>%
#'   print(target = "append_run.docx")
docx_append_run <- function( x, str, style = "Normal", pos = "after" ){

  style_id <- get_style_id(x=x, style=style, type = "character")
  xml_elt <- "<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:rPr><w:rStyle w:val=\"%s\"/></w:rPr><w:t xml:space=\"preserve\">%s</w:t></w:r>"
  xml_elt <- sprintf(xml_elt, style_id, str)
  add_xml_run(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title append an image
#' @description append an image into a paragraph of a docx object
#' @param x a docx device
#' @param src image filename
#' @param style text style
#' @param width height in inches
#' @param height height in inches
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @importFrom xml2 as_xml_document xml_find_first
#' @examples
#' library(magrittr)
#' img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
#' docx() %>%
#'   docx_add_par("R logo: ", style = "Normal") %>%
#'   docx_append_img(src = img.file, style = "strong", width = .3, height = .3) %>%
#'   print(target = "append_img.docx")
docx_append_img <- function( x, src, style = "Normal", width, height, pos = "after" ){

  style_id <- get_style_id(x=x, style=style, type = "character")

  ext_img <- external_img(src, width = width, height = height)
  xml_elt <- format(ext_img, type = "wml")
  xml_elt <- paste0(sprintf("<w:p %s>", base_ns), "<w:pPr/>", xml_elt, "</w:p>")

  rids <- docx_reference_img(x, src)
  xml_elt <- wml_link_images( xml_elt, rids )

  drawing_node <- as_xml_document(xml_elt) %>% xml_find_first("//w:r/w:drawing")

  wml_ <- "<w:r %s><w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>%s</w:r>"
  xml_elt <- sprintf(wml_, base_ns, style_id, as.character(drawing_node) )

  add_xml_run(x = x, str = xml_elt, pos = pos)
}


#' @export
#' @title add a wml string into a Word document
#' @importFrom xml2 as_xml_document xml_find_first xml_add_child
#' @description The function add a wml string into
#' the document after, before or on a cursor location.
#' @param x a docx object
#' @param str a wml string
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
add_xml_run <- function(x, str, pos){
  xml_elt <- as_xml_document(str)
  pos_list <- c("after", "before")

  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  cursor_elt <- xml_find_first(x$xml_doc, x$cursor)

  pos <- ifelse(pos=="after", length(xml_children(cursor_elt)), 1)
  cursor_elt <- xml_find_first(x$xml_doc, x$cursor)
  xml_add_child(.x = cursor_elt, .value = xml_elt, .where = pos )
  x
}
