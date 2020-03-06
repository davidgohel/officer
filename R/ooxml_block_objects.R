# caption ----

#' @export
#' @title caption block
#' @description Create a representation of a
#' caption that can be used for cross reference. The caption
#' can also be an auto numbered paragraph.
#' @param label a scalar character representing label to display
#' @param style paragraph style name
#' @param id cross reference identifier
#' @param autonum an object generated with function [run_autonum]
#' @examples
#'
#' @example examples/block_caption.R
#' @family block functions for reporting
block_caption <- function(label, style, id, autonum = NULL) {
  z <- list(
    label = label,
    id = id,
    autonum = autonum,
    style = style
  )
  class(z) <- c("block_caption", "block")

  z
}

#' @export
print.block_caption <- function(x, ...) {
  if (is.null(x$autonum)) {
    auton <- "[anto-num on]"
  } else {
    auton <- "[antonum off]"
  }
  cat("caption ", auton, ": ", x$label, "\n", sep = "")
}


#' @export
to_wml.block_caption <- function(x, add_ns = FALSE, base_document = NULL, ...) {
  if (is.null(base_document)) {
    base_document <- get_reference_value("docx")
  }

  if (is.character(base_document)) {
    base_document <- read_docx(path = base_document)
  } else if (!inherits(base_document, "rdocx")) {
    stop("base_document can only be the path to a docx file or an rdocx document.")
  }

  open_tag <- wp_ns_no
  if (add_ns) {
    open_tag <- wp_ns_yes
  }

  if (is.null(x$style)) {
    style <- base_document$default_styles$paragraph
  } else {
    style <- x$style
  }
  style_id <- get_style_id(data = base_document$styles, style = style, type = "paragraph")

  autonum <- ""
  if (!is.null(x$autonum)) {
    autonum <- to_wml(x$autonum)
  }

  run_str <- sprintf("<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", htmlEscapeCopy(x$label))
  run_str <- paste0(autonum, run_str)
  run_str <- bookmark(x$id, run_str)
  out <- sprintf(
    "%s<w:pPr><w:pStyle w:val=\"%s\"/></w:pPr>%s</w:p>",
    open_tag, style_id, run_str
  )

  out
}


# toc ----

#' @export
#' @title table of content
#' @description Create a representation of a table of content.
#' @param level max title level of the table
#' @param style optional. style in the document that will be used to build entries of the TOC.
#' @param separator optional. Some configurations need "," (i.e. from Canada) separator instead of ";"
#' @examples
#' block_toc(level = 2)
#' block_toc(style = "Table title")
#' @family block functions for reporting
block_toc <- function(level = 3, style = NULL, separator = ";") {
  z <- list(
    level = level, style = style, separator = separator
  )
  class(z) <- c("block_toc", "block")

  z
}

#' @export
print.block_toc <- function(x, ...) {
  if (is.null(x$style)) {
    cat("TOC - max level: ", x$level, "\n", sep = "")
  } else {
    cat("TOC for style: ", x$style, "\n", sep = "")
  }
}

#' @export
to_wml.block_toc <- function(x, add_ns = FALSE, ...) {

  open_tag <- wp_ns_no
  if (add_ns) {
    open_tag <- wp_ns_yes
  }


  if (is.null(x$style)) {
    out <- paste0(
      open_tag,
      "<w:pPr/>",
      to_wml(
        run_seqfield(
          seqfield = sprintf(
            "TOC \u005Co &quot;1-%.0f&quot; \u005Ch \u005Cz \u005Cu",
            x$level
          )
        )
      ),
      "</w:p>"
    )
  } else {
    out <- paste0(
      open_tag,
      "<w:pPr/>",
      to_wml(
        run_seqfield(
          seqfield = sprintf(
            "TOC \u005Ch \u005Cz \u005Ct \"%s%s1\"",
            x$style, x$separator
          )
        )
      ),
      "</w:p>"
    )
  }



  out
}



# section ----

#' @export
#' @title new section
#' @description Create a representation of a section
#' @param property section properties defined with function [prop_section]
#' @examples
#' prop_section(
#'   page_size = page_size(orient = "landscape"),
#'   page_margins = page_mar(top = 2),
#'   type = "continuous"
#' )
#' @family block functions for reporting
block_section <- function(property) {
  z <- list(
    property = property
  )
  class(z) <- c("block_section", "block")

  z
}

#' @export
print.block_section <- function(x, ...) {
  cat("----- end of section: ", "\n", sep = "")
}

#' @export
to_wml.block_section <- function(x, add_ns = FALSE, ...) {
  open_tag <- wp_ns_no
  if (add_ns) {
    open_tag <- wp_ns_yes
  }

  out <- paste0(open_tag,
    "<w:pPr>",
    to_wml(x$property),
    "</w:pPr></w:p>"
  )

  out
}


# table ----
table_docx <- function(x, header, style_id,
                       first_row, last_row, first_column,
                       last_column, no_hband, no_vband, add_ns = FALSE) {
  open_tag <- tbl_ns_no
  if (add_ns) {
    open_tag <- tbl_ns_yes
  }


  str <- paste0(open_tag,
    "<w:tblPr>",
    "<w:tblStyle w:val=\"", style_id, "\"/>",
    "<w:tblLook w:firstRow=\"", as.integer(first_row),
    "\" w:lastRow=\"", as.integer(last_row),
    "\" w:firstColumn=\"", as.integer(first_column),
    "\" w:lastColumn=\"", as.integer(last_column),
    "\" w:noHBand=\"", as.integer(no_hband),
    "\" w:noVBand=\"", as.integer(no_vband), "\"/>",
    "</w:tblPr>"
  )

  header_str <- character(length = 0L)
  if (header) {
    header_str <- paste0(
      "<w:tr><w:trPr><w:tblHeader/></w:trPr>",
      paste0("<w:tc><w:trPr/><w:p><w:r><w:t>",
        htmlEscapeCopy(enc2utf8(colnames(x))),
        "</w:t></w:r></w:p></w:tc>",
        collapse = ""
      ),
      "</w:tr>"
    )
  }

  as_tc <- function(x) {
    paste0(
      "<w:tc><w:trPr/><w:p><w:r><w:t>",
      htmlEscapeCopy(enc2utf8(x)),
      "</w:t></w:r></w:p></w:tc>"
    )
  }

  z <- lapply(x, as_tc)
  z <- do.call(paste0, z)
  z <- paste0("<w:tr>", z, "</w:tr>", collapse = "")

  paste0(str, header_str, z, "</w:tbl>")
}

table_pptx <- function(x, style_id, col_width, row_height,
                       first_row = TRUE, first_column = FALSE,
                       last_row = FALSE, last_column = FALSE, header = TRUE ){

  str <- paste0("<a:tbl>",
                sprintf("<a:tblPr firstRow=\"%.0f\" lastRow=\"%.0f\" firstColumn=\"%.0f\" lastColumn=\"%.0f\"",
                        first_row, last_row, first_column, last_column),
                ">",
                sprintf("<a:tableStyleId>%s</a:tableStyleId>", style_id),
                "</a:tblPr>",
                "<a:tblGrid>",
                paste0(sprintf("<a:gridCol w=\"%.0f\"/>", rep(col_width, length(x))), collapse = ""),
                "</a:tblGrid>")

  as_tc <- function(x) {
    paste0("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>",
           htmlEscapeCopy(enc2utf8(x)),
           "</a:t></a:r></a:p></a:txBody></a:tc>"
    )
  }
  header_str <- character(length = 0L)
  if( header ){
    header_str  <- paste0(
      sprintf("<a:tr h=\"%.0f\">", row_height),
      paste0(as_tc(colnames(x)), collapse = ""),
      "</a:tr>"
    )
  }

  z <- lapply(x, as_tc)
  z <- do.call(paste0, z)
  z <- paste0(sprintf("<a:tr h=\"%.0f\">", row_height), z, "</a:tr>", collapse = "")

  z <- paste0(str, header_str, z, "</a:tbl>")
  z <- paste0(
    "<a:graphic>",
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">",
    z, "</a:graphicData>", "</a:graphic>")
}


#' @export
#' @title table
#' @description Create a representation of a table
#' @param x a data.frame to add as a table
#' @param style table style
#' @param header display header if TRUE
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
#' block_table(x = mtcars)
#' @family block functions for reporting
block_table <- function(x, style = NULL, header = TRUE,
                        first_row = TRUE, first_column = FALSE,
                        last_row = FALSE, last_column = FALSE,
                        no_hband = FALSE, no_vband = TRUE) {

  stopifnot(is.data.frame(x))
  if(inherits(x, "tbl_df"))
    x <- as.data.frame(
      x, check.names = FALSE, stringsAsFactors = FALSE )

  z <- list(
    x = x,
    style = style,
    header = header,
    first_row = first_row,
    first_column = first_column,
    last_row = last_row,
    last_column = last_column,
    no_hband = no_hband,
    no_vband = no_vband
  )
  class(z) <- c("block_table", "block")

  z
}

#' @export
#' @importFrom utils str
print.block_table <- function(x, ...) {
  str(x$x)
}

#' @export
to_wml.block_table <- function(x, add_ns = FALSE, base_document = NULL, ...) {
  if (is.null(base_document)) {
    base_document <- get_reference_value("docx")
  }

  if (is.character(base_document)) {
    base_document <- read_docx(path = base_document)
  } else if (!inherits(base_document, "rdocx")) {
    stop("base_document can only be the path to a docx file or an rdocx document.")
  }

  if (is.null(x$style)) {
    style <- base_document$default_styles$paragraph
  } else {
    style <- x$style
  }
  style_id <- get_style_id(data = base_document$styles, style = style, type = "table")

  value <- characterise_df(x$x)

  out <- table_docx(
    x = value, header = x$header, style_id = style_id,
    first_row = x$first_row, last_row = x$last_row,
    first_column = x$first_column, last_column = x$last_column,
    no_hband = x$no_hband, no_vband = x$no_vband, add_ns = add_ns
  )

  out
}

#' @export
to_pml.block_table <- function(x, add_ns = FALSE,
                               left = 0, top = 0, width = 3, height = 3,
                               bg = "transparent", rot = 0, label = "", ph = "<p:ph/>", ...){

  if( !is.null(bg) && !is.color( bg ) )
    stop("bg must be a valid color.", call. = FALSE )

  bg_str <- solid_fill_pml(bg)

  xfrm_str <- p_xfrm_str(left = left, top = top, width = width, height = height, rot = rot)
  if( is.null(ph) || is.na(ph)){
    ph = "<p:ph/>"
  }

  value <- characterise_df(x$x)
  value_str <- table_pptx(value, style_id = x$style,
                        col_width = as.integer((width/ncol(x$x))*914400),
                        row_height = as.integer((height/nrow(x$x))*914400),
                        first_row = x$first_row, last_row = x$last_row,
                        first_column = x$first_column, last_column = x$last_column,
                        header = x$header )


  str <- paste0("<p:graphicFrame xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                "<p:nvGraphicFramePr>",
                sprintf("<p:cNvPr id=\"0\" name=\"%s\"/>", label),
                "<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr>",
                sprintf("<p:nvPr>%s</p:nvPr>", ph),
                "</p:nvGraphicFramePr>",
                xfrm_str,
                bg_str,
                value_str,
                "</p:graphicFrame>")
  str
}

# fpar ----

#' @export
#' @title concatenate formatted text as a paragraph
#' @description Create a paragraph representation by concatenating
#' formatted text or images.
#'
#' \code{fpar} supports \code{ftext}, \code{external_img} and simple strings.
#' All its arguments will be concatenated to create a paragraph where chunks
#' of text and images are associated with formatting properties.
#'
#' Default text and paragraph formatting properties can also
#' be modified with update.
#' @details
#' \code{fortify_fpar}, \code{as.data.frame} are used internally and
#' are not supposed to be used by end user.
#'
#' @param ... cot objects (ftext, external_img)
#' @param fp_p paragraph formatting properties
#' @param fp_t default text formatting properties. This is used as
#' text formatting properties when simple text is provided as argument.
#'
#' @param object fpar object
#' @examples
#' fpar(ftext("hello", shortcuts$fp_bold()))
#'
#' # mix text and image -----
#' img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
#'
#' bold_face <- shortcuts$fp_bold(font.size = 12)
#' bold_redface <- update(bold_face, color = "red")
#' fpar_1 <- fpar(
#'   "Hello World, ",
#'   ftext("how ", prop = bold_redface ),
#'   external_img(src = img.file, height = 1.06/2, width = 1.39/2),
#'   ftext(" you?", prop = bold_face ) )
#' fpar_1
#'
#' img_in_par <- fpar(
#'   external_img(src = img.file, height = 1.06/2, width = 1.39/2),
#'   fp_p = fp_par(text.align = "center") )
#' @family block functions for reporting
fpar <- function( ..., fp_p = fp_par(), fp_t = fp_text() ) {
  out <- list()
  out$chunks <- list(...)
  out$fp_p <- fp_p
  out$fp_t <- fp_t
  class(out) <- c("fpar")
  out
}

#' @export
#' @rdname fpar
#' @importFrom stats update
update.fpar <- function (object, fp_p = NULL, fp_t = NULL, ...){

  if(!is.null(fp_p)){
    object$fp_p <- fp_p
  }
  if(!is.null(fp_t)){
    object$fp_t <- fp_t
  }

  object
}


fortify_fpar <- function(x){
  lapply(x$chunks, function(chk) {
    if( !inherits(chk, c("cot", "run")) ){
      chk <- ftext(text = format(chk), prop = x$fp_t )
    }
    chk
  })
}


as.data.frame.fpar <- function( x, ...){
  chks <- fortify_fpar(x)
  chks <- chks[sapply(chks, function(x) inherits(x, "ftext"))]
  chks <- mapply(function(x){
    data.frame(value = x$value, size = x$pr$font.size,
               bold = x$pr$bold, italic = x$pr$italic,
               font.family = x$pr$font.family, stringsAsFactors = FALSE )
  }, chks, SIMPLIFY = FALSE)
  rbind.match.columns(chks)
}

#' @export
to_wml.fpar <- function(x, add_ns = FALSE, style_id = NULL, ...) {

  open_tag <- wp_ns_no
  if (add_ns) {
    open_tag <- wp_ns_yes
  }
  if(is.null(style_id)){
    par_style <- ppr_wml(x$fp_p)
  } else par_style <- paste0(
    "<w:pStyle w:val=\"", style_id, "\"/>")

  chks <- fortify_fpar(x)
  z <- lapply(chks, to_wml)
  z$collapse <- ""
  z <- do.call(paste0, z)
  paste0(open_tag, par_style, z, "</w:p>")
}

#' @export
to_pml.fpar <- function(x, add_ns = FALSE, ...) {

  open_tag <- ap_ns_no
  if (add_ns) {
    open_tag <- ap_ns_yes
  }

  par_style <- ppr_pml(x$fp_p)
  chks <- fortify_fpar(x)
  z <- lapply(chks, to_pml)
  z$collapse <- ""
  z <- do.call(paste0, z)
  paste0(open_tag, par_style, z, "</a:p>")
}


#' @export
to_html.fpar <- function(x, add_ns = FALSE, ...) {
  par_style <- ppr_css(x$fp_p)
  chks <- fortify_fpar(x)
  z <- lapply(chks, to_html)
  z$collapse <- ""
  z <- do.call(paste0, z)
  sprintf("<p style=\"%s\">%s</p>", par_style, z)
}

# block_list -----

#' @export
#' @title create paragraph blocks
#' @description a list of blocks can be used to gather
#' several blocks (paragraphs or tables) into a single
#' object. The function is to be used when adding
#' footnotes or formatted paragraphs into a new slide.
#' @param ... a list of objects of class \code{\link{fpar}} or
#' \code{flextable}. When output is only for Word, objects
#' of class \code{\link{external_img}} can also be used in
#' fpar construction to mix text and images in a single paragraph.
#' @examples
#'
#' @example examples/block_list.R
#' @seealso [ph_with()], [body_add()]
#' @family block functions for reporting
block_list <- function(...){
  x <- list(...)
  class(x) <- "block_list"
  x
}

# unordered list ----
#' @export
#' @title unordered list
#' @description unordered list of text for PowerPoint
#' presentations. Each text is associated with
#' a hierarchy level.
#' @param str_list list of strings to be included in the object
#' @param level_list list of levels for hierarchy structure
#' @param style text style, a \code{fp_text} object list or a
#' single \code{fp_text} objects. Use \code{fp_text(font.size = 0, ...)} to
#' inherit from default sizes of the presentation.
#' @examples
#' unordered_list(
#' level_list = c(1, 2, 2, 3, 3, 1),
#' str_list = c("Level1", "Level2", "Level2", "Level3", "Level3", "Level1"),
#' style = fp_text(color = "red", font.size = 0) )
#' unordered_list(
#' level_list = c(1, 2, 1),
#' str_list = c("Level1", "Level2", "Level1"),
#' style = list(
#'   fp_text(color = "red", font.size = 0),
#'   fp_text(color = "pink", font.size = 0),
#'   fp_text(color = "orange", font.size = 0)
#'   ))
#' @seealso \code{\link{ph_with}}
#' @family block functions for reporting
unordered_list <- function(str_list = character(0), level_list = integer(0), style = NULL){
  stopifnot(is.character(str_list))
  stopifnot(is.numeric(level_list))

  if (length(str_list) != length(level_list) & length(str_list) > 0) {
    stop("str_list and level_list have different lenghts.")
  }

  if( !is.null(style)){
    if( inherits(style, "fp_text") )
      style <- lapply(seq_len(length(str_list)), function(x) style )
  }
  x <- list(
    str = str_list,
    lvl = level_list,
    style = style
  )
  class(x) <- "unordered_list"
  x
}
#' @export
#' @noRd
print.unordered_list <- function(x, ...){
  print(data.frame(str = x$str,
                   lvl = x$lvl,
                   stringsAsFactors = FALSE))
  invisible()
}

#' @export
to_pml.unordered_list <- function(x, add_ns = FALSE, ...) {

  open_tag <- ap_ns_no
  if (add_ns) {
    open_tag <- ap_ns_yes
  }
  if( !is.null(x$style)){
    style_str <- sapply(x$style, format, type = "pml")
    style_str <- rep_len(style_str, length.out = length(x$str))
  } else style_str <- rep("<a:rPr/>", length(x$str))
  tmpl <- "%s<a:pPr%s/><a:r>%s<a:t>%s</a:t></a:r></a:p>"
  lvl <- sprintf(" lvl=\"%.0f\"", x$lvl - 1)
  lvl <- ifelse(x$lvl > 1, lvl, "")
  p <- sprintf(tmpl, open_tag, lvl, style_str, htmlEscapeCopy(x$str) )
  p <- paste(p, collapse = "")
  p
}

# plot_instr -----
#' @title Wrap plot instructions for png plotting in Powerpoint or Word
#' @description A simple wrapper to capture
#' plot instructions that will be executed and copied in a document. It produces
#' an object of class 'plot_instr' with a corresponding method [ph_with()].
#'
#' The function enable usage of any R plot with argument `code`. Wrap your code
#' between curly bracket if more than a single expression.
#'
#' @param code plotting instructions
#' @examples
#'
#' @example examples/plot_instr.R
#' @export
#' @import graphics
#' @seealso [ph_with()], [body_add()]
#' @family block functions for reporting
plot_instr <- function(code) {
  out <- list()
  out$code <- substitute(code)
  class(out) <- "plot_instr"
  return(out)
}

