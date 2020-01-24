htmlEscapeCopy <- local({

  .htmlSpecials <- list(
    `&` = '&amp;',
    `<` = '&lt;',
    `>` = '&gt;'
  )
  .htmlSpecialsPattern <- paste(names(.htmlSpecials), collapse='|')
  .htmlSpecialsAttrib <- c(
    .htmlSpecials,
    `'` = '&#39;',
    `"` = '&quot;',
    `\r` = '&#13;',
    `\n` = '&#10;'
  )
  .htmlSpecialsPatternAttrib <- paste(names(.htmlSpecialsAttrib), collapse='|')
  function(text, attribute=FALSE) {
    pattern <- if(attribute)
      .htmlSpecialsPatternAttrib
    else
      .htmlSpecialsPattern
    text <- enc2utf8(as.character(text))
    # Short circuit in the common case that there's nothing to escape
    if (!any(grepl(pattern, text, useBytes = TRUE)))
      return(text)
    specials <- if(attribute)
      .htmlSpecialsAttrib
    else
      .htmlSpecials
    for (chr in names(specials)) {
      text <- gsub(chr, specials[[chr]], text, fixed = TRUE, useBytes = TRUE)
    }
    Encoding(text) <- "UTF-8"
    return(text)
  }
})

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
#' block_caption("a caption", style = "Normal", id = "caption_id")
#' block_caption("a caption", style = "Normal", id = "caption_id",
#'   autonum = run_autonum())
#' @family block functions for reporting
block_caption <- function(label, style, id, autonum = NULL){

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
print.block_caption <- function(x, ...){
  if(is.null(x$autonum)){
    auton <- "[anto-num on]"
  } else {
    auton <- "[antonum off]"
  }
  cat("caption ", auton, ": ", x$label, "\n", sep = "")
}


#' @export
to_wml.block_caption = function (x, base_document = NULL, ...){

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
    else style <- x$style
    style_id <- get_style_id(data = base_document$styles, style=style, type = "paragraph")

    autonum <- ""
    if(!is.null(x$autonum)){
      autonum <- to_wml(x$autonum)
    }

    run_str <- sprintf( "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", htmlEscape(x$label))
    run_str <- paste0(autonum, run_str)
    run_str <- bookmark(x$id, run_str)
    out <- sprintf( "%s<w:pPr><w:pStyle w:val=\"%s\"/></w:pPr>%s</w:p>",
                    wml_with_ns("w:p"), style_id, run_str)

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
block_toc <- function(level = 3, style = NULL, separator = ";"){

  z <- list(
    level = level, style = style, separator = separator
  )
  class(z) <- c("block_toc", "block")

  z
}

#' @export
print.block_toc <- function(x, ...){
  if(is.null(x$style)){
    cat("TOC - max level: ", x$level, "\n", sep = "")
  } else {
    cat("TOC for style: ", x$style, "\n", sep = "")
  }

}

#' @export
to_wml.block_toc = function (x, ...){
    if(is.null(x$style)){
      out <- paste0(
        "<w:p><w:pPr/>",
        to_wml(
          run_seqfield(
            seqfield = sprintf("TOC \u005Co &quot;1-%.0f&quot; \u005Ch \u005Cz \u005Cu",
                               x$level))),
        "</w:p>")
    } else {
      out <- paste0(
        "<w:p><w:pPr/>",
        to_wml(
          run_seqfield(
            seqfield = sprintf("TOC \u005Ch \u005Cz \u005Ct \"%s%s1\"",
                               x$style, x$separator))),
        "</w:p>")
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
#'   type = "continuous")
#' @family block functions for reporting
block_section <- function(property){

  z <- list(
    property = property
  )
  class(z) <- c("block_section", "block")

  z
}

#' @export
print.block_section <- function(x, ...){
  cat("----- end of secion: ", "\n", sep = "")
}

#' @export
to_wml.block_section = function (x, ...){

    out <- paste0(
      "<w:p><w:pPr>",
      to_wml(x$property),
      "</w:pPr></w:p>")

  out
}


# table ----
table_docx <- function(x, header, style_id,
                       first_row, last_row, first_column,
                       last_column, no_hband, no_vband){
  str <- paste0(
    "<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
    " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"",
    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:tblPr>",
    "<w:tblStyle w:val=\"", style_id, "\"/>",
    "<w:tblLook w:firstRow=\"", as.integer(first_row),
    "\" w:lastRow=\"", as.integer(last_row),
    "\" w:firstColumn=\"", as.integer(first_column),
    "\" w:lastColumn=\"", as.integer(last_column),
    "\" w:noHBand=\"", as.integer(no_hband),
    "\" w:noVBand=\"", as.integer(no_vband), "\"/>",
    "</w:tblPr>")

  header_str <- character(length = 0L)
  if( header ){
    header_str  <- paste0(
      "<w:tr><w:trPr><w:tblHeader/></w:trPr>",
      paste0("<w:tc><w:trPr/><w:p><w:r><w:t>",
             htmlEscapeCopy(enc2utf8(colnames(x))),
             "</w:t></w:r></w:p></w:tc>", collapse = ""),
      "</w:tr>"
    )
  }

  as_tc <- function(x) {
    paste0("<w:tc><w:trPr/><w:p><w:r><w:t>",
           htmlEscapeCopy(enc2utf8(x)),
           "</w:t></w:r></w:p></w:tc>"
    )
  }

  z <- lapply(x, as_tc)
  z <- do.call(paste0, z)
  z <- paste0("<w:tr>", z, "</w:tr>", collapse = "")

  paste0(str, header_str, z, "</w:tbl>")
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
                        no_hband = FALSE, no_vband = TRUE){

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
# @importFrom utils str
print.block_table <- function(x, ...){
  str(x$x)
}

#' @export
to_wml.block_table = function (x, base_document = NULL, ...){
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
    else style <- x$style
    style_id <- get_style_id(data = base_document$styles, style=style, type = "table")

    value <- characterise_df(x$x)

    out <- table_docx(x = value, header = x$header, style_id = style_id,
                      first_row = x$first_row, last_row = x$last_row,
                      first_column = x$first_column, last_column = x$last_column,
                      no_hband = x$no_hband, no_vband = x$no_vband)

  out
}





