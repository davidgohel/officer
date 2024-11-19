#' @export
#' @title Add a page break in a 'Word' document
#' @description add a page break into an rdocx object
#' @param x an rdocx object
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' doc <- read_docx()
#' doc <- body_add_break(doc)
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_break <- function(x, pos = "after") {
  str <- runs_to_p_wml(run_pagebreak(), add_ns = TRUE)
  body_add_xml(x = x, str = str, pos = pos)
}

#' @export
#' @title Add an image in a 'Word' document
#' @description add an image into an rdocx object.
#' @inheritParams body_add_break
#' @param src image filename, the basename of the file must not contain any blank.
#' @param style paragraph style
#' @param width,height image size in units expressed by the unit argument.
#' Defaults to "in"ches.
#' @param unit One of the following units in which the width and height
#' arguments are expressed: "in", "cm" or "mm".
#' @examples
#' doc <- read_docx()
#'
#' img.file <- file.path(R.home("doc"), "html", "logo.jpg")
#' if (file.exists(img.file)) {
#'   doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39)
#'
#'   # Set the unit in which the width and height arguments are expressed
#'   doc <- body_add_img(
#'     x = doc, src = img.file,
#'     height = 2.69, width = 3.53,
#'     unit = "cm"
#'   )
#' }
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_img <- function(x, src, style = NULL, width, height, pos = "after", unit = "in") {
  if (is.null(style)) {
    style <- x$default_styles$paragraph
  }

  unit <- check_unit(unit, c("in", "cm", "mm"))

  file_type <- gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", src)

  if (file_type %in% ".svg") {
    if (!requireNamespace("rsvg")) {
      stop("package 'rsvg' is required to convert svg file to rasters.")
    }

    file <- tempfile(fileext = ".png")
    rsvg::rsvg_png(src, file = file)
    src <- file
    file_type <- ".png"
  }

  new_src <- tempfile(fileext = file_type)
  file.copy(src, to = new_src)

  style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")

  ext_img <- external_img(new_src, width = width, height = height, unit = unit)
  xml_elt <- runs_to_p_wml(ext_img, add_ns = TRUE, style_id = style_id)

  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title Add an external docx in a 'Word' document
#' @description Add content of a docx into an rdocx object.
#'
#' The function is using a 'Microsoft Word' feature: when the
#' document will be edited, the content of the file will be
#' inserted in the main document.
#'
#' This feature is unlikely to work as expected if the
#' resulting document is edited by another software.
#'
#' The file is added when the method `print()` that
#' produces the final Word file is called, so don't remove
#' file defined with `src` before.
#' @inheritParams body_add_break
#' @param src docx filename, the path of the file must not contain
#' any '&' and the basename must not contain any space.
#' @examples
#' file1 <- tempfile(fileext = ".docx")
#' file2 <- tempfile(fileext = ".docx")
#' file3 <- tempfile(fileext = ".docx")
#' x <- read_docx()
#' x <- body_add_par(x, "hello world 1", style = "Normal")
#' print(x, target = file1)
#'
#' x <- read_docx()
#' x <- body_add_par(x, "hello world 2", style = "Normal")
#' print(x, target = file2)
#'
#' x <- read_docx(path = file1)
#' x <- body_add_break(x)
#' x <- body_add_docx(x, src = file2)
#' print(x, target = file3)
#' @export
#' @family functions for adding content
body_add_docx <- function(x, src, pos = "after") {

  if(grepl(" ", basename(src))){
    stop("The basename of the file to be inserted cannot contain spaces.")
  }

  xml_elt <- to_wml(block_pour_docx(file = src), add_ns = TRUE)
  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title Add a 'ggplot' in a 'Word' document
#' @description add a ggplot as a png image into an rdocx object.
#' @inheritParams body_add_break
#' @param value ggplot object
#' @param style paragraph style
#' @param width,height plot size in units expressed by the unit argument.
#' Defaults to a width of 6 and a height of 5 "in"ches.
#' @param unit One of the following units in which the width and height
#' arguments are expressed: "in", "cm" or "mm".
#' @param res resolution of the png image in ppi
#' @param scale Multiplicative scaling factor, same as in ggsave
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param ... Arguments to be passed to png function.
#' @importFrom grDevices dev.off
#' @examples
#' if (require("ggplot2")) {
#'   doc <- read_docx()
#'
#'   gg_plot <- ggplot(data = iris) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length))
#'
#'   if (capabilities(what = "png")) {
#'     doc <- body_add_gg(doc, value = gg_plot, style = "centered")
#'
#'     # Set the unit in which the width and height arguments are expressed
#'     doc <- body_add_gg(doc, value = gg_plot, style = "centered", unit = "cm")
#'   }
#'
#'   print(doc, target = tempfile(fileext = ".docx"))
#' }
#' @family functions for adding content
#' @importFrom ragg agg_png
body_add_gg <- function(x, value, width = 6, height = 5, res = 300, style = "Normal", scale = 1, pos = "after", unit = "in", ...) {
  if (!requireNamespace("ggplot2")) {
    stop("package ggplot2 is required to use this function")
  }

  stopifnot(inherits(value, "gg"))

  if ("units" %in% names(list(...))) {
    cli::cli_abort(
      c("Found a {.arg units} argument. Did you mean {.arg unit}?")
    )
  }

  unit <- check_unit(unit, c("in", "cm", "mm"))


  file <- tempfile(fileext = ".png")
  agg_png(filename = file, width = width, height = height, scaling = scale, units = unit, res = res, background = "transparent", ...)
  print(value)
  dev.off()
  on.exit(unlink(file))
  body_add_img(x, src = file, style = style, width = width, height = height, pos = pos, unit = unit)
}


#' @export
#' @title Add a list of blocks into a 'Word' document
#' @description add a list of blocks produced by \code{block_list} into
#' into an rdocx object.
#' @inheritParams body_add_break
#' @param blocks set of blocks to be used as footnote content returned by
#'   function [block_list()].
#' @examples
#' library(officer)
#'
#' img.file <- file.path(R.home("doc"), "html", "logo.jpg")
#'
#' bl <- block_list(
#'   fpar(ftext("hello", shortcuts$fp_bold(color = "red"))),
#'   fpar(
#'     ftext("hello world", shortcuts$fp_bold()),
#'     external_img(src = img.file, height = 1.06, width = 1.39),
#'     fp_p = fp_par(text.align = "center")
#'   )
#' )
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_blocks(doc_1, blocks = bl)
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_blocks <- function(x, blocks, pos = "after") {
  stopifnot(inherits(blocks, "block_list"))

  if (length(blocks) > 0) {
    pos_vector <- rep("after", length(blocks))
    pos_vector[1] <- pos
    for (i in seq_along(blocks)) {
      x <- body_add_xml(x,
        str = to_wml(blocks[[i]], add_ns = TRUE),
        pos = pos_vector[i]
      )
    }
  }

  x
}




#' @export
#' @title Add paragraphs of text in a 'Word' document
#' @description add a paragraph of text into an rdocx object
#' @param x a docx device
#' @param value a character
#' @param style paragraph style name
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' doc <- read_docx()
#' doc <- body_add_par(doc, "A title", style = "heading 1")
#' doc <- body_add_par(doc, "Hello world!", style = "Normal")
#' doc <- body_add_par(doc, "centered text", style = "centered")
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_par <- function(x, value, style = NULL, pos = "after") {
  if (is.null(style)) {
    style <- x$default_styles$paragraph
  }

  style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")

  xml_elt <- paste0(
    wp_ns_yes,
    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr><w:r><w:t xml:space=\"preserve\">",
    htmlEscapeCopy(value), "</w:t></w:r></w:p>"
  )
  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title Add fpar in a 'Word' document
#' @description Add an \code{fpar} (a formatted paragraph)
#' into an rdocx object.
#' @param x a docx device
#' @param value a character
#' @param style paragraph style. If NULL, paragraph settings from `fpar` will be used. If not
#' NULL, it must be a paragraph style name (located in the template
#' provided as `read_docx(path = ...)`); in that case, paragraph settings from `fpar` will be
#' ignored.
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @examples
#' bold_face <- shortcuts$fp_bold(font.size = 30)
#' bold_redface <- update(bold_face, color = "red")
#' fpar_ <- fpar(
#'   ftext("Hello ", prop = bold_face),
#'   ftext("World", prop = bold_redface),
#'   ftext(", how are you?", prop = bold_face)
#' )
#' doc <- read_docx()
#' doc <- body_add_fpar(doc, fpar_)
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#'
#' # a way of using fpar to center an image in a Word doc ----
#' rlogo <- file.path(R.home("doc"), "html", "logo.jpg")
#' img_in_par <- fpar(
#'   external_img(src = rlogo, height = 1.06 / 2, width = 1.39 / 2),
#'   hyperlink_ftext(
#'     href = "https://cran.r-project.org/index.html",
#'     text = "cran", prop = bold_redface
#'   ),
#'   fp_p = fp_par(text.align = "center")
#' )
#'
#' doc <- read_docx()
#' doc <- body_add_fpar(doc, img_in_par)
#' print(doc, target = tempfile(fileext = ".docx"))
#'
#' @seealso \code{\link{fpar}}
#' @family functions for adding content
body_add_fpar <- function(x, value, style = NULL, pos = "after") {
  if (!is.null(style)) {
    style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")
  } else {
    style_id <- NULL
  }

  xml_elt <- to_wml(value, add_ns = TRUE, style_id = style_id)

  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title Add table in a 'Word' document
#' @description Add a table into an rdocx object.
#' @param x a docx device
#' @param value a data.frame to add as a table
#' @param style table style
#' @param pos where to add the new element relative to the cursor,
#' one of after", "before", "on".
#' @param header display header if TRUE
#' @param alignment columns alignement, argument length must match with columns length,
#' values must be "l" (left), "r" (right) or "c" (center).
#' @param align_table table alignment within document, value must be "left", "center" or "right"
#' @param stylenames columns styles defined by [table_stylenames()]
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
#' doc <- read_docx()
#' doc <- body_add_table(doc, iris, style = "table_template")
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_table <- function(x, value, style = NULL, pos = "after", header = TRUE,
                           alignment = NULL,
                           align_table = "center",
                           stylenames = table_stylenames(),
                           first_row = TRUE, first_column = FALSE,
                           last_row = FALSE, last_column = FALSE,
                           no_hband = FALSE, no_vband = TRUE) {
  pt <- prop_table(
    style = style, layout = table_layout(),
    width = table_width(), stylenames = stylenames,
    tcf = table_conditional_formatting(
      first_row = first_row, first_column = first_column,
      last_row = last_row, last_column = last_column,
      no_hband = no_hband, no_vband = no_vband
    ),
    align = align_table
  )

  bt <- block_table(x = value, header = header, properties = pt, alignment = alignment)
  xml_elt <- to_wml(bt, add_ns = TRUE, base_document = x)
  body_add_xml(x = x, str = xml_elt, pos = pos)
}

#' @export
#' @title Add table of content in a 'Word' document
#' @description Add a table of content into an rdocx object.
#' The TOC will be generated by Word, if the document is not
#' edited with Word (i.e. Libre Office) the TOC will not be generated.
#' @param x an rdocx object
#' @param level max title level of the table
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param style optional. style in the document that will be used to build entries of the TOC.
#' @param separator optional. Some configurations need "," (i.e. from Canada) separator instead of ";"
#' @examples
#' doc <- read_docx()
#' doc <- body_add_toc(doc)
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_toc <- function(x, level = 3, pos = "after", style = NULL, separator = ";") {
  bt <- block_toc(level = level, style = style, separator = separator)
  out <- to_wml(bt, add_ns = TRUE)
  body_add_xml(x = x, str = out, pos = pos)
}




#' @export
#' @title Add plot in a 'Word' document
#' @description Add a plot as a png image into an rdocx object.
#' @inheritParams body_add_break
#' @param value plot instructions, see [plot_instr()].
#' @param style paragraph style
#' @param width,height plot size in units expressed by the unit argument.
#' Defaults to a width of 6 and a height of 5 "in"ches.
#' @param unit One of the following units in which the width and height
#' arguments are expressed: "in", "cm" or "mm".
#' @param res resolution of the png image in ppi
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @param ... Arguments to be passed to png function.
#' @importFrom grDevices dev.off
#' @examples
#' doc <- read_docx()
#'
#' if (capabilities(what = "png")) {
#'   p <- plot_instr(
#'       code = {
#'         barplot(1:5, col = 2:6)
#'       }
#'     )
#'
#'   doc <- body_add_plot(doc, value = p, style = "centered")
#'
#'   # Set the unit in which the width and height arguments are expressed
#'   doc <- body_add_plot(doc, value = p, style = "centered", unit = "cm")
#' }
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family functions for adding content
body_add_plot <- function(x, value, width = 6, height = 5, res = 300, style = "Normal", pos = "after", unit = "in", ...) {
  unit <- check_unit(unit, c("in", "cm", "mm"))

  file <- tempfile(fileext = ".png")
  agg_png(filename = file, width = width, height = height, units = unit, res = res, background = "transparent", ...)
  tryCatch(
    {
      eval(value$code)
    },
    finally = {
      dev.off()
    }
  )
  on.exit(unlink(file))
  body_add_img(x, src = file, style = style, width = width, height = height, unit = unit, pos = pos)
}


#' @export
#' @title Add Word caption in a 'Word' document
#' @description Add a Word caption into an rdocx object.
#' @param x an rdocx object
#' @param value an object returned by [block_caption()]
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @family functions for adding content
#' @examples
#' doc <- read_docx()
#'
#' if (capabilities(what = "png")) {
#'   doc <- body_add_plot(doc,
#'     value = plot_instr(
#'       code = {
#'         barplot(1:5, col = 2:6)
#'       }
#'     ),
#'     style = "centered"
#'   )
#' }
#' run_num <- run_autonum(
#'   seq_id = "fig", pre_label = "Figure ",
#'   bkm = "barplot"
#' )
#' caption <- block_caption("a barplot",
#'   style = "Normal",
#'   autonum = run_num
#' )
#' doc <- body_add_caption(doc, caption)
#' print(doc, target = tempfile(fileext = ".docx"))
body_add_caption <- function(x, value, pos = "after") {
  stopifnot(inherits(value, "block_caption"))
  if (!value$style %in% x$styles$style_name) {
    stop("caption is using style ", shQuote(value$style), " that does not exist in the Word document.")
  }
  out <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml(x = x, str = out, pos = pos)
}




#' @export
#' @title Add an xml string as document element
#' @description Add an xml string as document element in the document. This function
#' is to be used to add custom openxml code.
#' @param x an rdocx object
#' @param str a wml string
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @keywords internal
body_add_xml <- function(x, str, pos = c("after", "before", "on")) {
  xml_elt <- as_xml_document(str)
  pos <- match.arg(pos)

  cursor_elt <- docx_current_block_xml(x)
  if (is.null(cursor_elt)) {
    xml_add_child(xml_find_first(x$doc_obj$get(), "/w:document/w:body"),
      xml_elt,
      .where = 0
    )
    .name <- xml_name(xml_elt)
    x$officer_cursor <- cursor_append(x$officer_cursor, .name)
  } else if (pos == "on") {
    xml_replace(cursor_elt, xml_elt)
    x$officer_cursor <- cursor_replace_nodename(x$officer_cursor, xml_name(xml_elt))
  } else if (pos == "after") {
    xml_add_sibling(cursor_elt, xml_elt, .where = pos)
    x$officer_cursor <- cursor_add_after(x$officer_cursor, xml_name(xml_elt))
  } else {
    xml_add_sibling(cursor_elt, xml_elt, .where = pos)
    x$officer_cursor <- cursor_add_before(x$officer_cursor, xml_name(xml_elt))
  }
  x
}

body_add_xml2 <- function(x, str) {
  x <- cursor_end(x)
  xml_elt <- as_xml_document(str)
  cursor_elt <- docx_current_block_xml(x)
  if (is.null(cursor_elt)) {
    xml_add_child(xml_find_first(x$doc_obj$get(), "/w:document/w:body"),
      xml_elt,
      .where = 0
    )
    x$officer_cursor <- cursor_append(x$officer_cursor, xml_name(xml_elt))
  } else {
    xml_add_sibling(cursor_elt, xml_elt, .where = "after")
    x$officer_cursor <- cursor_append(x$officer_cursor, xml_name(cursor_elt))
  }

  x
}

#' @export
#' @title Add bookmark in a 'Word' document
#' @description Add a bookmark at the cursor location. The bookmark
#' is added on the first run of text in the current paragraph.
#' @param x an rdocx object
#' @param id bookmark name
#' @examples
#'
#' # cursor_bookmark ----
#'
#' doc <- read_docx()
#' doc <- body_add_par(doc, "centered text", style = "centered")
#' doc <- body_bookmark(doc, "text_to_replace")
body_bookmark <- function(x, id) {
  cursor_elt <- docx_current_block_xml(x)
  ns_ <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
  new_id <- uuid_generate()
  id <- check_bookmark_id(id)

  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\" %s/>", new_id, id, ns_)
  bm_start_end <- sprintf("<w:bookmarkEnd %s w:id=\"%s\"/>", ns_, new_id)

  path_ <- paste0(xml_path(cursor_elt), "//w:r")

  node <- xml_find_first(x$doc_obj$get(), path_)
  xml_add_sibling(node, as_xml_document(bm_start_str), .where = "before")
  xml_add_sibling(node, as_xml_document(bm_start_end), .where = "after")

  x
}

#' @export
#' @title Remove an element in a 'Word' document
#' @description Remove element pointed by cursor from a
#' 'Word' document.
#' @param x an rdocx object
#' @examples
#' library(officer)
#'
#' str1 <- rep("Lorem ipsum dolor sit amet, consectetur adipiscing elit. ", 20)
#' str1 <- paste(str1, collapse = "")
#'
#' str2 <- "Drop that text"
#'
#' str3 <- rep("Aenean venenatis varius elit et fermentum vivamus vehicula. ", 20)
#' str3 <- paste(str3, collapse = "")
#'
#' my_doc <- read_docx()
#' my_doc <- body_add_par(my_doc, value = str1, style = "Normal")
#' my_doc <- body_add_par(my_doc, value = str2, style = "centered")
#' my_doc <- body_add_par(my_doc, value = str3, style = "Normal")
#'
#' new_doc_file <- print(my_doc,
#'   target = tempfile(fileext = ".docx")
#' )
#'
#' my_doc <- read_docx(path = new_doc_file)
#' my_doc <- cursor_reach(my_doc, keyword = "that text")
#' my_doc <- body_remove(my_doc)
#'
#' print(my_doc, target = tempfile(fileext = ".docx"))
body_remove <- function(x) {
  cursor_elt <- docx_current_block_xml(x)

  if (is.null(cursor_elt)) {
    warning("There is nothing left to remove in the document")
    return(x)
  }
  x$officer_cursor <- cursor_delete(x$officer_cursor, xml_name(cursor_elt))
  xml_remove(cursor_elt)
  x
}

#' @export
#' @title Add comment in a 'Word' document
#' @description Add a comment at the cursor location. The comment
#' is added on the first run of text in the current paragraph.
#' @param x an rdocx object
#' @param cmt a set of blocks to be used as comment content returned by
#'   function [block_list()].
#' @param author comment author.
#' @param date comment date
#' @param initials comment initials
#' @examples
#' doc <- read_docx()
#' doc <- body_add_par(doc, "Paragraph")
#' doc <- body_comment(doc, block_list("This is a comment."))
#' docx_file <- print(doc, target = tempfile(fileext = ".docx"))
#' docx_comments(read_docx(docx_file))
body_comment <- function(x,
                         cmt = ftext(""),
                         author = "",
                         date = "",
                         initials = "") {
  cursor_elt <- docx_current_block_xml(x)
  ns_ <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""

  open_tag <- wr_ns_yes

  blocks <- sapply(cmt, to_wml)
  blocks <- paste(blocks, collapse = "")

  id <- basename(tempfile(pattern = "comment"))

  cmt_xml <- paste0(
    sprintf(
      "<w:comment %s  w:id=\"%s\" w:author=\"%s\" w:date=\"%s\" w:initials=\"%s\">",
      ns_, id, author, date, initials
    ),
    blocks,
    "</w:comment>"
  )

  cmt_start_str <- sprintf("<w:commentRangeStart w:id=\"%s\" %s/>", id, ns_)
  cmt_start_end <- sprintf("<w:commentRangeEnd %s w:id=\"%s\"/>", ns_, id)

  path_ <- paste0(xml_path(cursor_elt), "//w:r")

  cmt_ref_xml <- paste0(
    open_tag,
    if (!is.null(x$pr)) rpr_wml(x$pr),
    "<w:commentReference w:id=\"",
    id,
    "\">",
    cmt_xml,
    "</w:commentReference>",
    "</w:r>"
  )

  path_ <- paste0(xml_path(cursor_elt), "//w:r")

  node <- xml_find_first(x$doc_obj$get(), path_)
  xml_add_sibling(node, as_xml_document(cmt_start_str), .where = "before")
  xml_add_sibling(node, as_xml_document(cmt_start_end), .where = "after")
  xml_add_sibling(node, as_xml_document(cmt_ref_xml), .where = "after")

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
#' doc_1 <- body_add(
#'   doc_1,
#'   block_section(prop_section(type = "continuous"))
#' )
#' doc_1 <- body_add(doc_1, plot_instr(code = barplot(1:5, col = 2:6)))
#' doc_1 <- body_add(
#'   doc_1,
#'   block_section(prop_section(page_size = page_size(orient = "landscape")))
#' )
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' # print(doc_1, target = "test.docx")
#' @section Illustrations:
#'
#' \if{html}{\figure{body_add_doc_1.png}{options: width=70\%}}
#' @keywords internal
body_add <- function(x, value, ...) {
  UseMethod("body_add", value)
}


#' @export
#' @describeIn body_add add a character vector.
#' @param style paragraph style name. These names are available with function [styles_info]
#' and are the names of the Word styles defined in the base document (see
#' argument `path` from [read_docx]).
body_add.character <- function(x, value, style = NULL, ...) {
  if (!is.null(style)) {
    style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")
  } else {
    style_id <- NULL
  }

  runs <- paste0(
    "<w:r><w:t xml:space=\"preserve\">",
    htmlEscapeCopy(value), "</w:t></w:r>"
  )

  xml_elt <- paste0(wp_ns_yes, "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>", runs, "</w:p>")
  for (str in xml_elt) {
    x <- body_add_xml2(x = x, str = str)
  }
  x
}

#' @export
#' @describeIn body_add add a numeric vector.
#' @param format_fun a function to be used to format values.
body_add.numeric <- function(x, value, style = NULL, format_fun = formatC, ...) {
  value <- format_fun(value, ...)
  body_add(x, value = value, style = style, ...)
}

#' @export
#' @describeIn body_add add a factor vector.
body_add.factor <- function(x, value, style = NULL, format_fun = as.character, ...) {
  value <- format_fun(value)
  body_add(x, value = value, style = style, ...)
}



#' @export
#' @describeIn body_add add a [fpar] object. These objects enable
#' the creation of formatted paragraphs made of formatted chunks of text.
body_add.fpar <- function(x, value, style = NULL, ...) {
  img_src <- sapply(value$chunks, function(x) {
    if (inherits(x, "external_img")) {
      as.character(x)
    } else {
      NA_character_
    }
  })
  img_src <- unique(img_src[!is.na(img_src)])

  if (!is.null(style)) {
    style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")
  } else {
    style_id <- NULL
  }

  xml_elt <- to_wml(value, add_ns = TRUE, style_id = style_id)

  body_add_xml2(x = x, str = xml_elt)
}

#' @export
#' @param header display header if TRUE
#' @param tcf conditional formatting settings defined by [table_conditional_formatting()]
#' @param alignment columns alignement, argument length must match with columns length,
#' values must be "l" (left), "r" (right) or "c" (center).
#' @describeIn body_add add a data.frame object with [block_table()].
body_add.data.frame <- function(x, value, style = NULL, header = TRUE,
                                tcf = table_conditional_formatting(),
                                alignment = NULL,
                                ...) {
  pt <- prop_table(
    style = style, layout = table_layout(),
    width = table_width(),
    tcf = tcf
  )

  bt <- block_table(x = value, header = header, properties = pt, alignment = alignment)
  xml_elt <- to_wml(bt, add_ns = TRUE, base_document = x)

  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add a [block_caption] object. These objects enable
#' the creation of set of formatted paragraphs made of formatted chunks of text.
body_add.block_caption <- function(x, value, ...) {

  if (!value$style %in% x$styles$style_name) {
    stop("caption is using style ", shQuote(value$style), " that does not exist in the Word document.")
  }

  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add a [block_list] object.
body_add.block_list <- function(x, value, ...) {
  if (length(value) > 0) {
    for (i in seq_along(value)) {
      x <- body_add(x, value = value[[i]])
    }
  }

  x
}

#' @export
#' @describeIn body_add add a table of content (a [block_toc] object).
body_add.block_toc <- function(x, value, ...) {
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add an image (a [external_img] object).
body_add.external_img <- function(x, value, style = "Normal", ...) {
  if (is.null(style)) {
    style <- x$default_styles$paragraph
  }

  file_type <- gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", value)

  if (file_type %in% ".svg") {
    if (!requireNamespace("rsvg")) {
      stop("package 'rsvg' is required to convert svg file to rasters.")
    }

    file <- tempfile(fileext = ".png")
    rsvg::rsvg_png(src, file = file)
    value <- external_img(
      src = file,
      width = attr(value, "dims")$width,
      height = attr(value, "dims")$height
    )
    file_type <- ".png"
  }

  new_src <- tempfile(fileext = file_type)
  file.copy(value, to = new_src)
  value[1] <- new_src

  style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")

  xml_elt <- paste0(
    wp_ns_yes,
    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
    to_wml(value, add_ns = FALSE),
    "</w:p>"
  )

  body_add_xml2(x = x, str = xml_elt)
}

#' @export
#' @describeIn body_add add a [run_pagebreak] object.
body_add.run_pagebreak <- function(x, value, style = NULL, ...) {
  if (is.null(style)) {
    style <- x$default_styles$paragraph
  }

  style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")

  xml_elt <- paste0(
    wp_ns_yes,
    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
    to_wml(value, add_ns = FALSE),
    "</w:p>"
  )

  body_add_xml2(x = x, str = xml_elt)
}

#' @export
#' @describeIn body_add add a [run_columnbreak] object.
body_add.run_columnbreak <- function(x, value, style = NULL, ...) {
  if (is.null(style)) {
    style <- x$default_styles$paragraph
  }

  style_id <- get_style_id(data = x$styles, style = style, type = "paragraph")

  xml_elt <- paste0(
    wp_ns_yes,
    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>",
    to_wml(value, add_ns = FALSE),
    "</w:p>"
  )

  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add add a ggplot object.
#' @param width,height plot size in units expressed by the unit argument.
#' Defaults to a width of 6 and a height of 5 "in"ches.
#' @param unit One of the following units in which the width and height
#' arguments are expressed: "in", "cm" or "mm".
#' @param res resolution of the png image in ppi
#' @param scale Multiplicative scaling factor, same as in ggsave
body_add.gg <- function(x, value, width = 6, height = 5, res = 300, style = "Normal", scale = 1, unit = "in", ...) {
  if (!requireNamespace("ggplot2")) {
    stop("package ggplot2 is required to use this function")
  }

  unit <- check_unit(unit, c("in", "cm", "mm"))

  file <- tempfile(fileext = ".png")
  agg_png(filename = file, width = width, height = height, units = unit, res = res, scaling = scale, background = "transparent", ...)
  print(value)
  dev.off()
  on.exit(unlink(file))

  value <- external_img(src = file, width = width, height = height, unit = unit)
  body_add(x, value, style = style)
}

#' @export
#' @describeIn body_add add a base plot with a [plot_instr] object.
body_add.plot_instr <- function(x, value, width = 6, height = 5, res = 300, style = "Normal", unit = "in", ...) {
  unit <- check_unit(unit, c("in", "cm", "mm"))

  file <- tempfile(fileext = ".png")
  agg_png(filename = file, width = width, height = height, units = unit, res = res, scaling = 1, background = "transparent", ...)
  tryCatch(
    {
      eval(value$code)
    },
    finally = {
      dev.off()
    }
  )
  on.exit(unlink(file))

  value <- external_img(src = file, width = width, height = height, unit = unit)

  body_add(x, value, style = style)
}

#' @export
#' @describeIn body_add pour content of an external docx file with with a [block_pour_docx] object
body_add.block_pour_docx <- function(x, value, ...) {
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml2(x = x, str = xml_elt)
}


#' @export
#' @describeIn body_add ends a section with a [block_section] object
body_add.block_section <- function(x, value, ...) {
  xml_elt <- to_wml(value, add_ns = TRUE, base_document = x)
  body_add_xml2(x = x, str = xml_elt)
}
