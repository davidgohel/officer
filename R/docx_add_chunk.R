#' @export
#' @title append seq field
#' @description append seq field into a paragraph of an rdocx object
#' @param x an rdocx object
#' @param str seq field value
#' @param style text style
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @examples
#' library(magrittr)
#' x <- read_docx() %>%
#'   body_add_par("Time is: ", style = "Normal") %>%
#'   slip_in_seqfield(
#'     str = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT",
#'     style = 'strong') %>%
#'
#'   body_add_par(" - This is a figure title", style = "centered") %>%
#'   slip_in_seqfield(str = "SEQ Figure \u005C* roman",
#'     style = 'Default Paragraph Font', pos = "before") %>%
#'   slip_in_text("Figure: ", style = "strong", pos = "before") %>%
#'
#'   body_add_par(" - This is another figure title", style = "centered") %>%
#'   slip_in_seqfield(str = "SEQ Figure \u005C* roman",
#'     style = 'strong', pos = "before")  %>%
#'   slip_in_text("Figure: ", style = "strong", pos = "before") %>%
#'   body_add_par("This is a symbol: ", style = "Normal") %>%
#'   slip_in_seqfield(str = "SYMBOL 100 \u005Cf Wingdings",
#'     style = 'strong')
#'
#' print(x, target = tempfile(fileext = ".docx"))
slip_in_seqfield <- function( x, str, style = NULL, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$character

  style_id <- get_style_id(data = x$styles, style=style, type = "character")

  xml_elt_1 <- paste0(wml_with_ns("w:r"),
                      sprintf("<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>", style_id),
                      "<w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/>",
                      "</w:r>")
  xml_elt_2 <- paste0(wml_with_ns("w:r"),
                      sprintf("<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>", style_id),
                      sprintf("<w:instrText xml:space=\"preserve\" w:dirty=\"true\">%s</w:instrText>", str ),
                      "</w:r>")
  xml_elt_3 <- paste0(wml_with_ns("w:r"),
                      sprintf("<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>", style_id),
                      "<w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/>",
                      "</w:r>")

  if( pos == "after"){
    slip_in_xml(x = x, str = xml_elt_1, pos = pos)
    slip_in_xml(x = x, str = xml_elt_2, pos = pos)
    slip_in_xml(x = x, str = xml_elt_3, pos = pos)
  } else {
    slip_in_xml(x = x, str = xml_elt_3, pos = pos)
    slip_in_xml(x = x, str = xml_elt_2, pos = pos)
    slip_in_xml(x = x, str = xml_elt_1, pos = pos)
  }

}

#' @export
#' @title append text
#' @description append text into a paragraph of an rdocx object
#' @param x an rdocx object
#' @param str text
#' @param style text style
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @examples
#' library(magrittr)
#' x <- read_docx() %>%
#'   body_add_par("Hello ", style = "Normal") %>%
#'   slip_in_text("world", style = "strong") %>%
#'   slip_in_text("Message is", style = "strong", pos = "before")
#'
#' print(x, target = tempfile(fileext = ".docx"))
slip_in_text <- function( x, str, style = NULL, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$character

  style_id <- get_style_id(data = x$styles, style=style, type = "character")
  xml_elt <- paste0( wml_with_ns("w:r"),
      "<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>",
      "<w:t xml:space=\"preserve\">%s</w:t></w:r>")
  xml_elt <- sprintf(xml_elt, style_id, str)
  slip_in_xml(x = x, str = xml_elt, pos = pos)
}




#' @export
#' @title append an image
#' @description append an image into a paragraph of an rdocx object
#' @param x an rdocx object
#' @param src image filename, the basename of the file must not contain any blank.
#' @param style text style
#' @param width height in inches
#' @param height height in inches
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @importFrom xml2 as_xml_document xml_find_first
#' @examples
#' library(magrittr)
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' x <- read_docx() %>%
#'   body_add_par("R logo: ", style = "Normal") %>%
#'   slip_in_img(src = img.file, style = "strong", width = .3, height = .3)
#'
#' print(x, target = tempfile(fileext = ".docx"))
slip_in_img <- function( x, src, style = NULL, width, height, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$character

  new_src <- tempfile( fileext = gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", src) )
  file.copy( src, to = new_src )

  style_id <- get_style_id(data = x$styles, style=style, type = "character")

  ext_img <- external_img(new_src, width = width, height = height)
  xml_elt <- format(ext_img, type = "wml")
  xml_elt <- paste0(wml_with_ns("w:p"), "<w:pPr/>", xml_elt, "</w:p>")

  x <- docx_reference_img(x, new_src)
  xml_elt <- wml_link_images( x, xml_elt )

  drawing_node <- xml_find_first(as_xml_document(xml_elt), "//w:r/w:drawing")

  wml_ <- paste0(wml_with_ns("w:r"), "<w:rPr><w:rStyle w:val=\"%s\"/></w:rPr>%s</w:r>")
  xml_elt <- sprintf(wml_, style_id, as.character(drawing_node) )

  slip_in_xml(x = x, str = xml_elt, pos = pos)
}


#' @export
#' @title add a wml string into a Word document
#' @importFrom xml2 as_xml_document xml_find_first xml_add_child
#' @description The function add a wml string into
#' the document after, before or on a cursor location.
#' @param x an rdocx object
#' @param str a wml string
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
slip_in_xml <- function(x, str, pos){
  xml_elt <- as_xml_document(str)
  pos_list <- c("after", "before")

  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  cursor_elt <- x$doc_obj$get_at_cursor()
  pos <- ifelse(pos=="after", length(xml_children(cursor_elt)), 1)
  xml_add_child(.x = cursor_elt, .value = xml_elt, .where = pos )
  x
}
