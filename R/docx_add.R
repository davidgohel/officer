#' @export
#' @title add page break
#' @description add a page break into an rdocx object
#' @param x an rdocx object
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' library(magrittr)
#' doc <- read_docx() %>% body_add_break()
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_break <- function( x, pos = "after"){
  str <- runs_to_p_wml(run_pagebreak(), add_ns = TRUE)
  body_add_xml(x = x, str = str, pos = pos)
}

#' @export
#' @title add image
#' @description add an image into an rdocx object.
#' @inheritParams body_add_break
#' @param src image filename, the basename of the file must not contain any blank.
#' @param style paragraph style
#' @param width height in inches
#' @param height height in inches
#' @examples
#' doc <- read_docx()
#'
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
#' }
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_img <- function( x, src, style = NULL, width, height, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$paragraph

  file_type <- gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", src)

  if( file_type %in% ".svg" ){
    if (!requireNamespace("rsvg")){
      stop("package 'rsvg' is required to convert svg file to rasters")
    }

    file <- tempfile(fileext = ".png")
    rsvg::rsvg_png(src, file = file)
    src <- file
    file_type <- ".png"
  }

  new_src <- tempfile( fileext = file_type )
  file.copy( src, to = new_src )

  style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")

  ext_img <- external_img(new_src, width = width, height = height)
  xml_elt <- runs_to_p_wml(ext_img, add_ns = TRUE, style_id = style_id)
  x <- docx_reference_img(x, new_src)
  xml_elt <- wml_link_images( x, xml_elt )

  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title insert an external docx
#' @description add content of a docx into an rdocx object.
#' @note
#' The function is using a 'Microsoft Word' feature: when the
#' document will be edited, the content of the file will be
#' inserted in the main document.
#'
#' This feature is unlikely to work as expected if the
#' resulting document is edited by another software.
#' @inheritParams body_add_break
#' @param src docx filename
#' @examples
#'
#' library(magrittr)
#' file1 <- tempfile(fileext = ".docx")
#' file2 <- tempfile(fileext = ".docx")
#' file3 <- tempfile(fileext = ".docx")
#' read_docx() %>%
#'   body_add_par("hello world 1", style = "Normal") %>%
#'   print(target = file1)
#' read_docx() %>%
#'   body_add_par("hello world 2", style = "Normal") %>%
#'   print(target = file2)
#'
#' read_docx(path = file1) %>%
#'   body_add_break() %>%
#'   body_add_docx(src = file2) %>%
#'   print(target = file3)
#' @export
#' @family functions for adding content
body_add_docx <- function( x, src, pos = "after" ){
  src <- unique( src )
  rel <- x$doc_obj$relationship()
  new_rid <- sprintf("rId%.0f", rel$get_next_id())
  new_docx_file <- basename(tempfile(fileext = ".docx"))
  file.copy(src, to = file.path(x$package_dir, new_docx_file))
  rel$add(
    id = new_rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
    target = file.path("../", new_docx_file) )
  x$content_type$add_override(
    setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", paste0("/", new_docx_file) )
  )
  x$content_type$save()
  xml_elt <- paste0("<w:altChunk xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
                    "r:id=\"", new_rid, "\"/>")
  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add ggplot
#' @description add a ggplot as a png image into an rdocx object
#' @inheritParams body_add_break
#' @param value ggplot object
#' @param style paragraph style
#' @param width height in inches
#' @param height height in inches
#' @param res resolution of the png image in ppi
#' @param ... Arguments to be passed to png function.
#' @importFrom grDevices png dev.off
#' @examples
#' if( require("ggplot2") ){
#'   doc <- read_docx()
#'
#'   gg_plot <- ggplot(data = iris ) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length))
#'
#'   if( capabilities(what = "png") )
#'     doc <- body_add_gg(doc, value = gg_plot, style = "centered" )
#'
#'   print(doc, target = tempfile(fileext = ".docx") )
#' }
#' @family functions for adding content
body_add_gg <- function( x, value, width = 6, height = 5, res = 300, style = "Normal", ... ){

  if( !requireNamespace("ggplot2") )
    stop("package ggplot2 is required to use this function")

  stopifnot(inherits(value, "gg") )
  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')
  png(filename = file, width = width, height = height, units = "in", res = res, ...)
  print(value)
  dev.off()
  on.exit(unlink(file))
  body_add_img(x, src = file, style = style, width = width, height = height)
}


#' @export
#' @title add a list of blocks into a document
#' @description add a list of blocks produced by \code{block_list} into
#' into an rdocx object
#' @inheritParams body_add_break
#' @param blocks set of blocks to be used as footnote content returned by
#'   function \code{\link{block_list}}.
#' @examples
#' library(magrittr)
#'
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#' bl <- block_list(
#'   fpar(ftext("hello", shortcuts$fp_bold())),
#'   fpar(
#'     ftext("hello world", shortcuts$fp_bold()),
#'     external_img(src = img.file, height = 1.06, width = 1.39)
#'   )
#' )
#'
#' x <- read_docx() %>%
#'   body_add_blocks( blocks = bl ) %>%
#'   print(target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_blocks <- function( x, blocks, pos = "after" ){
  stopifnot(inherits(blocks, "block_list"))

  if( length(blocks) > 0 ){
    pos_vector <- rep("after", length(blocks))
    pos_vector[1] <- pos
    for(i in seq_along(blocks) ){
      x <- body_add_fpar(x, value = blocks[[i]], pos = pos_vector[i])
    }
  }

  x
}




#' @export
#' @title add paragraph of text
#' @description add a paragraph of text into an rdocx object
#' @param x a docx device
#' @param value a character
#' @param style paragraph style name
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_par("A title", style = "heading 1") %>%
#'   body_add_par("Hello world!", style = "Normal") %>%
#'   body_add_par("centered text", style = "centered")
#'
#' print(doc, target = tempfile(fileext = ".docx") )
#' @family functions for adding content
body_add_par <- function( x, value, style = NULL, pos = "after" ){

  if( is.null(style) )
    style <- x$default_styles$paragraph

  style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")

  xml_elt <- paste0(wp_ns_yes,
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr><w:r><w:t xml:space=\"preserve\">",
                    htmlEscapeCopy(value), "</w:t></w:r></w:p>")
  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add fpar
#' @description add an \code{fpar} (a formatted paragraph) into an rdocx object
#' @param x a docx device
#' @param value a character
#' @param style paragraph style. If NULL, paragraph settings from `fpar` will be used. If not
#' NULL, it must be a paragraph style name (located in the template
#' provided as `read_docx(path = ...)`); in that case, paragraph settings from `fpar` will be
#' ignored.
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' library(magrittr)
#' bold_face <- shortcuts$fp_bold(font.size = 30)
#' bold_redface <- update(bold_face, color = "red")
#' fpar_ <- fpar(ftext("Hello ", prop = bold_face),
#'               ftext("World", prop = bold_redface ),
#'               ftext(", how are you?", prop = bold_face ) )
#' doc <- read_docx() %>% body_add_fpar(fpar_)
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#'
#' # a way of using fpar to center an image in a Word doc ----
#' rlogo <- file.path( R.home("doc"), "html", "logo.jpg" )
#' img_in_par <- fpar(
#'   external_img(src = rlogo, height = 1.06/2, width = 1.39/2),
#'   fp_p = fp_par(text.align = "center") )
#'
#' read_docx() %>% body_add_fpar(img_in_par) %>%
#'   print(target = tempfile(fileext = ".docx") )
#'
#' @seealso \code{\link{fpar}}
#' @family functions for adding content
body_add_fpar <- function( x, value, style = NULL, pos = "after" ){

  img_src <- sapply(value$chunks, function(x){
    if( inherits(x, "external_img"))
      as.character(x)
    else NA_character_
  })
  img_src <- unique(img_src[!is.na(img_src)])

  if( !is.null(style) ){
    style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")
  } else style_id <- NULL

  xml_elt <- to_wml(value, add_ns = TRUE, style_id = style_id)
  x <- docx_reference_img(x, img_src)
  xml_elt <- wml_link_images( x, xml_elt )

  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add table
#' @description add a table into an rdocx object
#' @param x a docx device
#' @param value a data.frame to add as a table
#' @param style table style
#' @param pos where to add the new element relative to the cursor,
#' one of after", "before", "on".
#' @param header display header if TRUE
#' @param alignment columns alignement, argument length must match with columns length,
#' values must be "l" (left), "r" (right) or "c" (center).
#' @param first_row Specifies that the first column conditional formatting should be
#' applied. Details for this and other conditional formatting options can be found
#' at http://officeopenxml.com/WPtblLook.php.
#' @param last_row Specifies that the first column conditional formatting should be applied.
#' @param first_column Specifies that the first column conditional formatting should
#' be applied.
#' @param last_column Specifies that the first column conditional formatting should be
#' applied.
#' @param no_hband Specifies that the first column conditional formatting should be applied.
#' @param no_vband Specifies that the first column conditional formatting should be applied.
#' @examples
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_table(iris, style = "table_template")
#'
#' print(doc, target = tempfile(fileext = ".docx") )
#' @family functions for adding content
body_add_table <- function( x, value, style = NULL, pos = "after", header = TRUE,
                            alignment = NULL,
                            first_row = TRUE, first_column = FALSE,
                            last_row = FALSE, last_column = FALSE,
                            no_hband = FALSE, no_vband = TRUE ){

  pt <- prop_table(
    style = style, layout = table_layout(),
    width = table_width(),
    tcf = table_conditional_formatting(
      first_row = first_row, first_column = first_column,
      last_row = last_row, last_column = last_column,
      no_hband = no_hband, no_vband = no_vband))

  bt <- block_table(x = value, header = header, properties = pt, alignment = alignment)
  xml_elt <- to_wml(bt, add_ns = TRUE, base_document = x)
  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title add table of content
#' @description add a table of content into an rdocx object.
#' The TOC will be generated by Word, if the document is not
#' edited with Word (i.e. Libre Office) the TOC will not be generated.
#' @param x an rdocx object
#' @param level max title level of the table
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param style optional. style in the document that will be used to build entries of the TOC.
#' @param separator optional. Some configurations need "," (i.e. from Canada) separator instead of ";"
#' @examples
#' library(magrittr)
#' doc <- read_docx() %>% body_add_toc()
#'
#' print(doc, target = tempfile(fileext = ".docx") )
#' @family functions for adding content
body_add_toc <- function( x, level = 3, pos = "after", style = NULL, separator = ";"){
  bt <- block_toc(level = level, style = style, separator = separator)
  out <- to_wml(bt, add_ns = TRUE)
  body_add_xml(x = x, str = out, pos = pos)
}

#' @export
#' @title add an xml string as document element
#' @description Add an xml string as document element in the document. This function
#' is to be used to add custom openxml code.
#' @param x an rdocx object
#' @param str a wml string
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
body_add_xml <- function(x, str, pos){
  xml_elt <- as_xml_document(str)
  pos_list <- c("after", "before", "on")

  if( !pos %in% pos_list )
    stop("unknown pos ", shQuote(pos, type = "sh"), ", it should be ",
         paste( shQuote(pos_list, type = "sh"), collapse = " or ") )

  cursor_elt <- x$doc_obj$get_at_cursor()

  if( length(x) == 1 ){
    xml_add_sibling(cursor_elt, xml_elt, .where = "before")
    x <- cursor_end(x)
  } else {
    if( pos != "on")
      xml_add_sibling(cursor_elt, xml_elt, .where = pos)
    else {
      xml_replace(cursor_elt, xml_elt)
    }
    if(pos == "after")
      x <- cursor_forward(x)
  }


  x
}

body_add_xml2 <- function(x, str){
  xml_elt <- as_xml_document(str)
  last_elt <- xml_find_first(x$doc_obj$get(), "w:body/*[last()]")
  xml_add_sibling(last_elt, xml_elt, .where = "before")
  cursor_end(x)
}

#' @export
#' @importFrom uuid UUIDgenerate
#' @title add bookmark
#' @description Add a bookmark at the cursor location. The bookmark
#' is added on the first run of text in the current paragraph.
#' @param x an rdocx object
#' @param id bookmark name
#' @examples
#'
#' # cursor_bookmark ----
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_par("centered text", style = "centered") %>%
#'   body_bookmark("text_to_replace")
body_bookmark <- function(x, id){
  cursor_elt <- x$doc_obj$get_at_cursor()
  ns_ <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
  new_id <- UUIDgenerate()

  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\" %s/>", new_id, id, ns_ )
  bm_start_end <- sprintf("<w:bookmarkEnd %s w:id=\"%s\"/>", ns_, new_id )

  path_ <- paste0( xml_path(cursor_elt), "//w:r")

  node <- xml_find_first(x$doc_obj$get(), path_ )
  xml_add_sibling(node, as_xml_document( bm_start_str ) , .where = "before")
  xml_add_sibling(node, as_xml_document( bm_start_end ) , .where = "after")

  x
}

#' @export
#' @title remove an element
#' @description remove element pointed by cursor from a Word document
#' @param x an rdocx object
#' @examples
#' library(officer)
#' library(magrittr)
#'
#' str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>%
#'   rep(20) %>% paste(collapse = "")
#' str2 <- "Drop that text"
#' str3 <- "Aenean venenatis varius elit et fermentum vivamus vehicula. " %>%
#'   rep(20) %>% paste(collapse = "")
#'
#' my_doc <- read_docx()  %>%
#'   body_add_par(value = str1, style = "Normal") %>%
#'   body_add_par(value = str2, style = "centered") %>%
#'   body_add_par(value = str3, style = "Normal")
#'
#' new_doc_file <- print(my_doc,
#'   target = tempfile(fileext = ".docx"))
#'
#' my_doc <- read_docx(path = new_doc_file)  %>%
#'   cursor_reach(keyword = "that text") %>%
#'   body_remove()
#'
#' print(my_doc, target = tempfile(fileext = ".docx"))
body_remove <- function(x){

  cursor_elt <- x$doc_obj$get_at_cursor()

  if( x$doc_obj$length() == 1 && xml_name(cursor_elt) == "sectPr"){
    warning("There is nothing left to remove in the document")
    return(x)
  }

  x$doc_obj$cursor_forward()
  new_cursor_elt <- x$doc_obj$get_at_cursor()
  xml_remove(cursor_elt)
  x$doc_obj$set_cursor(xml_path(new_cursor_elt))
  x
}


# body_add and methods -----
#' @export
#' @title Add content into a Word document
#' @description This function add objects into a Word document. Values are added
#' as new paragraphs or tables.
#'
#' This function is experimental and will replace the `body_add_*` functions
#' later. For now it is only to be used for successive additions and cannot
#' be used in conjunction with the `body_add_*` functions.
#' @param x an rdocx object
#' @param value object to add in the document. Supported objects
#' are vectors, data.frame, graphics, block of formatted paragraphs,
#' unordered list of formatted paragraphs,
#' pretty tables with package flextable, 'Microsoft' charts with package mschart.
#' @param ... further arguments passed to or from other methods. When
#' adding a `ggplot` object or [plot_instr], these arguments will be used
#' by png function. See method signatures to see what arguments can be used.
#' @examples
#' doc_1 <- read_docx()
#' doc_1 <- body_add(doc_1, "Table of content", style = "heading 1")
#' doc_1 <- body_add(doc_1, block_toc())
#' doc_1 <- body_add(doc_1, run_pagebreak())
#' doc_1 <- body_add(doc_1, "A title", style = "heading 1")
#' doc_1 <- body_add(doc_1, head(iris), style = "table_template")
#' doc_1 <- body_add(doc_1, "Another title", style = "heading 1")
#' doc_1 <- body_add(doc_1, letters, style = "Normal")
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' # print(doc_1, target = "test.docx")
#' @section Illustrations:
#'
#' \if{html}{\figure{body_add_doc_1.png}{options: width=60\%}}
body_add <- function(x, value, ...){
  UseMethod("body_add", value)
}


#' @export
#' @describeIn body_add add a character vector.
#' @param style paragraph style name. These names are available with function [styles_info]
#' and are the names of the Word styles defined in the base document (see
#' argument `path` from [read_docx]).
body_add.character <- function(x, value, style = NULL, ...){

  if( !is.null(style) ){
    style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")
  } else style_id <- NULL

  runs <- paste0("<w:r><w:t xml:space=\"preserve\">",
                 htmlEscapeCopy(value), "</w:t></w:r>")

  xml_elt <- paste0(wp_ns_yes, "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>", runs, "</w:p>")
  for(str in xml_elt){
    x <- body_add_xml2(x = x, str = str)
  }
  x
}

#' @export
#' @describeIn body_add add a numeric vector.
#' @param format_fun a function to be used to format values.
body_add.numeric <- function(x, value, style = NULL, format_fun = formatC, ...){
  value <- format_fun(value, ...)
  body_add(x, value = value, style = style, ...)
}

#' @export
#' @describeIn body_add add a factor vector.
body_add.factor <- function(x, value, style = NULL, format_fun = as.character, ...){
  value <- format_fun(value)
  body_add(x, value = value, style = style, ...)
}



#' @export
#' @describeIn body_add add a [fpar] object. These objects enable
#' the creation of formatted paragraphs made of formatted chunks of text.
body_add.fpar <- function( x, value, style = NULL, ... ){
  img_src <- sapply(value$chunks, function(x){
    if( inherits(x, "external_img"))
      as.character(x)
    else NA_character_
  })
  img_src <- unique(img_src[!is.na(img_src)])

  if( !is.null(style) ){
    style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")
  } else style_id <- NULL

  xml_elt <- to_wml(value, add_ns = TRUE, style_id = style_id)
  x <- docx_reference_img(x, img_src)
  xml_elt <- wml_link_images( x, xml_elt )
  body_add_xml2(x = x, str = xml_elt)
}

#' @export
#' @param header display header if TRUE
#' @param tcf conditional formatting settings defined by [table_conditional_formatting()]
#' @param alignment columns alignement, argument length must match with columns length,
#' values must be "l" (left), "r" (right) or "c" (center).
#' @describeIn body_add add a data.frame object with [block_table()].
body_add.data.frame <- function( x, value, style = NULL, header = TRUE,
                                 tcf = table_conditional_formatting(),
                                 alignment = NULL,
                                 ... ){

  pt <- prop_table(
    style = style, layout = table_layout(),
    width = table_width(),
    tcf = tcf)

  bt <- block_table(x = value, header = header, properties = pt, alignment = alignment)
  xml_elt <- to_wml(bt, add_ns = TRUE, base_document = x)

  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add a [block_caption] object. These objects enable
#' the creation of set of formatted paragraphs made of formatted chunks of text.
body_add.block_caption <- function( x, value, ... ){
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add a [block_list] object.
body_add.block_list <- function( x, value, ... ){

  if( length(value) > 0 ){
    for(i in seq_along(value) ){
      x <- body_add(x, value = value[[i]])
    }
  }

  x
}

#' @export
#' @describeIn body_add add a table of content (a [block_toc] object).
body_add.block_toc <- function( x, value, ... ){
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add an image (a [external_img] object).
body_add.external_img <- function( x, value, style = "Normal", ... ){

  if( is.null(style) )
    style <- x$default_styles$paragraph

  file_type <- gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", value)

  if( file_type %in% ".svg" ){
    if (!requireNamespace("rsvg")){
      stop("package 'rsvg' is required to convert svg file to rasters")
    }

    file <- tempfile(fileext = ".png")
    rsvg::rsvg_png(src, file = file)
    value <- external_img(src = file,
                          width = attr(value, "dims")$width,
                          height = attr(value, "dims")$height)
    file_type <- ".png"
  }

  new_src <- tempfile( fileext = file_type )
  file.copy( value, to = new_src )
  value[1] <- new_src

  style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")

  xml_elt <- paste0(wp_ns_yes,
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
                    to_wml(value, add_ns = FALSE),
                    "</w:p>")

  x <- docx_reference_img(x, new_src)
  xml_elt <- wml_link_images( x, xml_elt )

  body_add_xml2(x = x, str = xml_elt)
}

#' @export
#' @describeIn body_add add a [run_pagebreak] object.
body_add.run_pagebreak <- function( x, value, style = NULL, ... ){

  if( is.null(style) )
    style <- x$default_styles$paragraph

  style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")

  xml_elt <- paste0(wp_ns_yes,
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
                    to_wml(value, add_ns = FALSE),
                    "</w:p>")

  body_add_xml2(x = x, str = xml_elt)
}

#' @export
#' @describeIn body_add add a [run_columnbreak] object.
body_add.run_columnbreak <- function( x, value, style = NULL, ... ){

  if( is.null(style) )
    style <- x$default_styles$paragraph

  style_id <- get_style_id(data = x$styles, style=style, type = "paragraph")

  xml_elt <- paste0(wp_ns_yes,
                    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
                    to_wml(value, add_ns = FALSE),
                    "</w:p>")

  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add a ggplot object.
#' @param width height in inches
#' @param height height in inches
#' @param res resolution of the png image in ppi
body_add.gg <- function( x, value, width = 6, height = 5, res = 300, style = "Normal", ... ){

  if( !requireNamespace("ggplot2") )
    stop("package ggplot2 is required to use this function")

  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')
  png(filename = file, width = width, height = height, units = "in", res = res, ...)
  print(value)
  dev.off()
  on.exit(unlink(file))

  value <- external_img(src = file, width = width, height = height)
  body_add(x, value, style = style)
}

#' @export
#' @describeIn body_add add a base plot with a [plot_instr] object.
body_add.plot_instr <- function( x, value, width = 6, height = 5, res = 300, style = "Normal", ... ){

  file <- tempfile(fileext = ".png")
  options(bitmapType='cairo')
  png(filename = file, width = width, height = height, units = "in", res = res, ...)
  tryCatch({
    eval(value$code)
  },
  finally = {
    dev.off()
  } )
  on.exit(unlink(file))

  value <- external_img(src = file, width = width, height = height)

  body_add(x, value, style = style)
}


#' @export
#' @describeIn body_add pour content of an external docx file with with a [block_pour_docx] object
body_add.block_pour_docx <- function( x, value, ... ){
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml2(x = x, str = xml_elt)
}

