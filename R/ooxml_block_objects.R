# caption ----

#' @export
#' @title Caption block
#' @description Create a representation of a
#' caption that can be used for cross reference.
#' @param label a scalar character representing label to display
#' @param style paragraph style name
#' @param autonum an object generated with function [run_autonum]
#' @examples
#'
#' @example examples/block_caption.R
#' @family block functions for reporting
block_caption <- function(label, style, autonum = NULL) {
  z <- list(
    label = label,
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

to_wml_block_caption_officer <- function(x, add_ns = FALSE, base_document = NULL){
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

  out <- sprintf(
    "%s<w:pPr><w:pStyle w:val=\"%s\"/></w:pPr>%s</w:p>",
    open_tag, style_id, run_str
  )

  out
}
to_wml_block_caption_pandoc <- function(x, bookdown_id = NULL){

  if(is.null(x$label)) return("")

  autonum <- ""
  if (!is.null(x$autonum)) {
    autonum <- paste("`", to_wml(x$autonum), "`{=openxml}", sep = "")
  }

  run_str <- paste0(autonum, htmlEscapeCopy(x$label))

  paste0(
    if (!is.null(x$style)) paste0("\n\n::: {custom-style=\"", x$style, "\"}"),
    "\n\n",
    # "<caption>\n\n",
    if (!is.null(bookdown_id)) bookdown_id,
    run_str,
    # "\n\n</caption>",
    if (!is.null(x$style)) paste0("\n:::\n"),
    "\n\n"
  )
}

#' @export
to_wml.block_caption <- function(x, add_ns = FALSE, base_document = NULL, knitting = FALSE, ...) {
  if(knitting)
    to_wml_block_caption_pandoc(x, bookdown_id = list(...)$bookdown_id)
  else
    to_wml_block_caption_officer(x, add_ns = add_ns, base_document = base_document)
}


# toc ----

#' @export
#' @title Table of content
#' @description Create a representation of a table of content.
#' @param level max title level of the table
#' @param style optional. If not NULL, its value is used as style in the
#' document that will be used to build entries of the TOC.
#' @param seq_id optional. If not NULL, its value is used as sequence
#' identifier in the document that will be used to build entries of the
#' TOC. See also [run_autonum()] to specify a sequence identifier.
#' @param separator optional. Some configurations need "," (i.e. from Canada) separator instead of ";"
#' @examples
#' block_toc(level = 2)
#' block_toc(style = "Table Caption")
#' @family block functions for reporting
block_toc <- function(level = 3, style = NULL, seq_id = NULL, separator = ";") {
  z <- list(
    level = level, style = style, seq_id = seq_id, separator = separator
  )
  class(z) <- c("block_toc", "block")

  z
}

#' @export
print.block_toc <- function(x, ...) {
  if (is.null(x$style) && is.null(x$seq_id)) {
    cat("TOC - max level: ", x$level, "\n", sep = "")
  } else if (!is.null(x$style)) {
    cat("TOC for style: ", x$style, "\n", sep = "")
  } else if (!is.null(x$seq_id)) {
    cat("TOC for seq identifier: ", x$seq_id, "\n", sep = "")
  }
}

#' @export
to_wml.block_toc <- function(x, add_ns = FALSE, ...) {

  open_tag <- wp_ns_no
  if (add_ns) {
    open_tag <- wp_ns_yes
  }


  if(is.null(x$style) && is.null(x$seq_id)) {
    out <- paste0(
      open_tag,
      "<w:pPr/>",
      to_wml(
        run_seqfield(
          seqfield = sprintf(
            "TOC \\o \"1-%.0f\" \\h \\z \\u",
            x$level
          )
        )
      ),
      "</w:p>"
    )
  } else if(!is.null(x$style)) {
    out <- paste0(
      open_tag,
      "<w:pPr/>",
      to_wml(
        run_seqfield(
          seqfield = sprintf(
            "TOC \\h \\z \\t \"%s%s1\"",
            x$style, x$separator
          )
        )
      ),
      "</w:p>"
    )
  } else if(!is.null(x$seq_id)) {
    out <- paste0(
      open_tag,
      "<w:pPr/>",
      to_wml(
        run_seqfield(
          seqfield = sprintf("TOC \\h \\z \\c \"%s\"", x$seq_id)
        )
      ),
      "</w:p>"
    )
  }



  out
}

# pour_docx ----

#' @export
#' @title Pour external Word document in the current document
#' @description Pour the content of a docx file in the resulting docx
#' generated by the main R Markdown document.
#' @param file external docx file path
#' @examples
#' library(officer)
#' docx <- tempfile(fileext = ".docx")
#' doc <- read_docx()
#' doc <- body_add(doc, iris[1:20,], style = "table_template")
#' print(doc, target = docx)
#'
#' target <- tempfile(fileext = ".docx")
#' doc_1 <- read_docx()
#' doc_1 <- body_add(doc_1, block_pour_docx(docx))
#' print(doc_1, target = target)
#' @family block functions for reporting
block_pour_docx <- function(file){
  if(!file.exists(file)){
    stop("file {", file, "} does not exist.", call. = FALSE)
  }
  if(!grepl("\\.docx$", file, ignore.case = TRUE)){
    stop("file {", file, "} is not a docx file.", call. = FALSE)
  }
  z <- list(file = file)
  class(z) <- c("block_pour_docx", "block")
  z
}

#' @export
print.block_pour_docx <- function(x, ...) {
  cat("Pour docx: ", x$file, "\n", sep = "")
}

#' @export
to_wml.block_pour_docx <- function(x, add_ns = FALSE, base_document = NULL, ...) {
  if (is.null(base_document)) {
    base_document <- get_reference_value("docx")
  }

  if (is.character(base_document)) {
    base_document <- read_docx(path = base_document)
  } else if (!inherits(base_document, "rdocx")) {
    stop("base_document can only be the path to a docx file or an rdocx document.")
  }

  rel <- base_document$doc_obj$relationship()
  new_rid <- sprintf("rId%.0f", rel$get_next_id())
  new_docx_file <- basename(tempfile(fileext = ".docx"))
  file.copy(x$file, to = file.path(base_document$package_dir, new_docx_file))
  rel$add(
    id = new_rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
    target = file.path("../", new_docx_file) )
  base_document$content_type$add_override(
    setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", paste0("/", new_docx_file) )
  )
  base_document$content_type$save()
  out <- paste0("<w:altChunk xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
                    "r:id=\"", new_rid, "\"/>")


  out
}



# section ----

#' @export
#' @title New Word section
#' @description Create a representation of a section.
#'
#' A section affects preceding paragraphs or tables; i.e.
#' a section starts at the end of the previous section (or the beginning of
#' the document if no preceding section exists), and stops where the
#' section is declared.
#'
#' When a new landscape section is needed, it is recommended to add a block_section
#' with `type = "continuous"`, to add the content to be appened in the new section
#' and finally to add a block_section with `page_size = page_size(orient = "landscape")`.
#' @param property section properties defined with function [prop_section]
#' @examples
#' ps <- prop_section(
#'   page_size = page_size(orient = "landscape"),
#'   page_margins = page_mar(top = 2),
#'   type = "continuous"
#' )
#' block_section(ps)
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


# table properties ----


#' @export
#' @title Table conditional formatting
#' @description Tables can be conditionally formatted based on few properties as
#' whether the content is in the first row, last row, first column, or last
#' column, or whether the rows or columns are to be banded.
#' @param first_row,last_row apply or remove formatting from the first or last row in the table.
#' @param first_column,last_column apply or remove formatting from the first or last column in the table.
#' @param no_hband,no_vband don't display odd and even rows or columns with
#' alternating shading for ease of reading.
#' @note
#' You must define a format for first_row, first_column and other properties
#' if you need to use them. The format is defined in a docx template.
#' @examples
#' table_conditional_formatting(first_row = TRUE, first_column = TRUE)
#' @family functions for table definition
table_conditional_formatting <- function(
  first_row = TRUE, first_column = FALSE,
  last_row = FALSE, last_column = FALSE,
  no_hband = FALSE, no_vband = TRUE){

  z <- list(first_row = first_row, first_column = first_column,
            last_row = last_row, last_column = last_column,
            no_hband = no_hband, no_vband = no_vband)
  class(z) <- c("table_conditional_formatting")
  z
}

#' @export
to_wml.table_conditional_formatting <- function(x, add_ns = FALSE, ...) {
  paste0("<w:tblLook w:firstRow=\"", as.integer(x$first_row),
         "\" w:lastRow=\"", as.integer(x$last_row),
         "\" w:firstColumn=\"", as.integer(x$first_column),
         "\" w:lastColumn=\"", as.integer(x$last_column),
         "\" w:noHBand=\"", as.integer(x$no_hband),
         "\" w:noVBand=\"", as.integer(x$no_vband), "\"/>")
}

#' @export
to_pml.table_conditional_formatting <- function(x, add_ns = FALSE, ...){
  expr_ <- paste0(" firstRow=\"%.0f\" lastRow=\"%.0f\"",
         " firstColumn=\"%.0f\" lastColumn=\"%.0f\"",
         " bandRow=\"%.0f\" bandCol=\"%.0f\""
         )
  sprintf(expr_,
          x$first_row, x$last_row,
          x$first_column, x$last_column,
          !x$no_hband, !x$no_vband)
}

table_layout_types <- c("autofit", "fixed")


#' @export
#' @title Algorithm for table layout
#' @description When a table is displayed in a document, it can
#' either be displayed using a fixed width or autofit layout algorithm:
#'
#' * fixed: uses fixed widths for columns. The width of the table is not
#' changed regardless of the contents of the cells.
#' * autofit: uses the contents of each cell and the table width to
#' determine the final column widths.
#' @param type 'autofit' or 'fixed' algorithm. Default to 'autofit'.
#' @family functions for table definition
table_layout <- function(type = "autofit"){

  if(!type %in% table_layout_types){
    stop("type must be one of ", paste(table_layout_types, collapse = ", "), ".")
  }

  z <- list(type = type)
  class(z) <- "table_layout"
  z
}

#' @export
to_wml.table_layout <- function(x, add_ns = FALSE, ...) {
  sprintf("<w:tblLayout w:type=\"%s\"/>", x$type)
}

table_layout_width_units <- c("in", "pct")

#' @export
#' @title Preferred width for a table
#' @description Define the preferred width for a table.
#' @section Word:
#' All widths in a table are considered preferred because widths of
#' columns can conflict and the table layout rules can require a
#' preference to be overridden.
#' @param width value of the preferred width of the table.
#' @param unit unit of the width. Possible values are 'in' (inches) and 'pct' (percent)
#' @family functions for table definition
table_width <- function(width = 1, unit = "pct"){
  if(!unit %in% table_layout_width_units){
    stop("unit must be one of ", paste(table_layout_width_units, collapse = ", "), ".")
  }

  z <- list(width = width, unit = unit)
  class(z) <- "table_width"
  z

}
#' @export
to_wml.table_width <- function(x, add_ns = FALSE, ...) {
  if(x$unit %in% "pct"){
    tbl_width <- sprintf("<w:tblW w:type=\"pct\" w:w=\"%.0f\"/>",
                         x$width * 5000)
  } else {
    tbl_width <- sprintf("<w:tblW w:type=\"dxa\" w:w=\"%.0f\"/>",
                         x$width * 1440)
  }
  tbl_width
}

#' @export
#' @title Column widths of a table
#' @description The function defines the size of each column of a table.
#' @param widths Column widths expressed in inches.
#' @family functions for table definition
table_colwidths <- function(widths = NULL){
  z <- list(widths = widths)
  class(z) <- "table_colwidths"
  z
}

#' @export
#' @title Paragraph styles for columns
#' @description The function defines the paragraph styles for columns.
#' @param stylenames a named character vector, names are column names, values are
#' paragraph styles associated with each column. If a column is not
#' specified, default value 'Normal' is used.
#' Another form is as a named list, the list names are the styles
#' and the contents are column names to be formatted with the
#' corresponding style.
#' @family functions for table definition
#' @examples
#' library(officer)
#'
#' stylenames <- c(
#'   vs = "centered", am = "centered",
#'   gear = "centered", carb = "centered"
#' )
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_table(doc_1,
#'   value = mtcars, style = "table_template",
#'   stylenames = table_stylenames(stylenames = stylenames)
#' )
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#'
#'
#' stylenames <- list(
#'   "centered" = c("vs", "am", "gear", "carb")
#' )
#'
#' doc_2 <- read_docx()
#' doc_2 <- body_add_table(doc_2,
#'   value = mtcars, style = "table_template",
#'   stylenames = table_stylenames(stylenames = stylenames)
#' )
#'
#' print(doc_2, target = tempfile(fileext = ".docx"))
table_stylenames <- function(stylenames = list()){

  if( length(stylenames) > 0 && is.null(attr(stylenames, "names")) ){
    stop("stylenames should have names")
  }

  if( length(stylenames) > 0 && is.list(stylenames) ){
    .l <- vapply(stylenames, length, FUN.VALUE = 0L)
    zz <- inverse.rle(
      structure(list(
        lengths = .l,
        values = names(stylenames)),
        class = "rle")
      )
    names(zz) <- as.character(unlist(stylenames))
    stylenames <- as.list(zz)
  } else if(is.character(stylenames)){
    stylenames <- as.list(stylenames)
  }

  z <- list(stylenames = stylenames)
  class(z) <- "table_stylenames"
  z
}

#' @export
#' @importFrom utils modifyList
to_wml.table_stylenames <- function(x, add_ns = FALSE, base_document = NULL, dat, ...) {

  if (is.null(base_document)) {
    base_document <- get_reference_value("docx")
  }
  if (is.character(base_document)) {
    base_document <- read_docx(path = base_document)
  } else if (!inherits(base_document, "rdocx")) {
    stop("base_document can only be the path to a docx file or an rdocx document.")
  }

  stylenames <- rep("Normal", ncol(dat))
  names(stylenames) <- colnames(dat)
  stylenames <- as.list(stylenames)
  # restrict to only existing cols
  x$stylenames <- x$stylenames[names(x$stylenames) %in% colnames(dat)]
  stylenames <- modifyList(stylenames, val = x$stylenames)
  stylenames <- lapply(stylenames, function(x){
    get_style_id(data = base_document$styles, style = x, type = "paragraph")
  })
  stylenames <- lapply(stylenames, function(x){
    sprintf("<w:pStyle w:val=\"%s\"/>", x)
  })
  stylenames
}

#' @export
to_wml.table_colwidths <- function(x, add_ns = FALSE, ...) {
  if(length(x$widths) < 1) return("")
  grid_col_str <- sprintf("<w:gridCol w:w=\"%.0f\"/>", x$widths * 1440)
  grid_col_str <- paste(grid_col_str, collapse = "")
  paste0("<w:tblGrid>", grid_col_str, "</w:tblGrid>")
}

#' @export
#' @title Table properties
#' @description Define table properties such as fixed or autofit layout,
#' table width in the document, eventually column widths.
#' @param style table style to be used to format table
#' @param layout layout defined by [table_layout()],
#' @param width table width in the document defined by [table_width()]
#' @param stylenames columns styles defined by [table_stylenames()]
#' @param colwidths column widths defined by [table_colwidths()]
#' @param align table alignment (one of left, center or right)
#' @param tcf conditional formatting settings defined by [table_conditional_formatting()]
#' @examples
#' prop_table()
#' to_wml(prop_table())
#' @family functions for table definition
prop_table <- function(style = NA_character_, layout = table_layout(),
                       width = table_width(),
                       stylenames = table_stylenames(),
                       colwidths = table_colwidths(),
                       tcf = table_conditional_formatting(),
                       align = "center"){


  z <- list(
    style = style,
    layout = layout,
    width = width,
    colsizes = colwidths,
    stylenames = stylenames,
    tcf = tcf, align = align
  )
  class(z) <- c("prop_table")
  z
}

#' @export
to_wml.prop_table <- function(x, add_ns = FALSE, base_document = NULL, ...) {

  if (is.null(base_document)) {
    base_document <- get_reference_value("docx")
  }
  if (is.character(base_document)) {
    base_document <- read_docx(path = base_document)
  } else if (!inherits(base_document, "rdocx")) {
    stop("base_document can only be the path to a docx file or an rdocx document.")
  }

  style_id <- NA_character_
  if(!is.null(x$style) && !is.na(x$style)){
    if (is.null(x$style)) {
      style <- base_document$default_styles$table
    } else {
      style <- x$style
    }
    style_id <- get_style_id(data = base_document$styles, style = style, type = "table")
  }

  tbl_layout <- to_wml(x$layout, add_ns= add_ns, base_document = base_document)


  width <- ""
  if(!is.null(x$width) && "autofit" %in% x$type)
    width <- to_wml(x$width, add_ns= add_ns, base_document = base_document)

  colwidths <- to_wml(x$colsizes, add_ns= add_ns, base_document = base_document)
  tcf <- to_wml(x$tcf, add_ns= add_ns, base_document = base_document)
  paste0("<w:tblPr>",
         if(!is.na(style_id)) paste0("<w:tblStyle w:val=\"", style_id, "\"/>"),
         tbl_layout,
         sprintf( "<w:jc w:val=\"%s\"/>", x$align ),
         width, tcf,
         "</w:tblPr>",
         if(x$layout$type %in% "fixed") colwidths
         )

}


# table ----
table_docx <- function(x, header, style_id,
                       properties, alignment = NULL, add_ns = FALSE,
                       base_document = base_document) {
  open_tag <- tbl_ns_no
  if (add_ns) {
    open_tag <- tbl_ns_yes
  }

  str <- paste0(
    open_tag,
    to_wml(properties, add_ns = add_ns, base_document = base_document)
  )

  stylenames <- to_wml(properties$stylenames, base_document = base_document, dat = x)
  stylenames <- unlist(stylenames)
  stylenames <- as.character(stylenames)

  if(is.null(alignment)){
    alignment <- rep("", ncol(x))
  } else{
    alignment <- match.arg(alignment, c("left", "right", "center"), several.ok = TRUE )
    if(length(alignment) < ncol(x)){
      alignment <- rep(alignment, length.out = ncol(x) )
    }
    alignment <- sprintf("<w:jc w:val=\"%s\"/>", alignment)
  }

  header_str <- character(length = 0L)
  if (header) {

    header_str <- paste0(
      "<w:tr><w:trPr><w:tblHeader/></w:trPr>",
      paste0("<w:tc><w:p>",
             sprintf("<w:pPr>%s%s</w:pPr>", stylenames, alignment),
             "<w:r><w:t>",
        htmlEscapeCopy(enc2utf8(colnames(x))),
        "</w:t></w:r></w:p></w:tc>",
        collapse = ""
      ),
      "</w:tr>"
    )
  }
  as_tc <- function(x, align, stylenames) {
    paste0(
      "<w:tc><w:p>",
      sprintf("<w:pPr>%s%s</w:pPr>", stylenames, align),
      "<w:r><w:t>",
      htmlEscapeCopy(enc2utf8(x)),
      "</w:t></w:r></w:p></w:tc>"
    )
  }
  z <- mapply(as_tc, x, alignment, stylenames, SIMPLIFY = FALSE)
  z <- do.call(paste0, z)
  z <- paste0("<w:tr>", z, "</w:tr>", collapse = "")

  paste0(str, header_str, z, "</w:tbl>")
}

table_pptx <- function(x, style_id, col_width, row_height,
                       tcf, header = TRUE, alignment = NULL ){
  str <- paste0("<a:tbl>",
                sprintf("<a:tblPr %s>", to_pml(tcf)),
                sprintf("<a:tableStyleId>%s</a:tableStyleId>", style_id),
                "</a:tblPr>",
                "<a:tblGrid>",
                paste0(sprintf("<a:gridCol w=\"%.0f\"/>",
                               rep(col_width, length(x))),
                       collapse = ""),
                "</a:tblGrid>")

  as_tc <- function(x, align) {
    paste0("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>",
           "<a:pPr algn=\"", align, "\"/>",
           "<a:r><a:t>",
           htmlEscapeCopy(enc2utf8(x)),
           "</a:t></a:r></a:p></a:txBody></a:tc>"
    )
  }

  if(is.null(alignment)){
    alignment <- rep("r", ncol(x))
  } else{
    alignment <- match.arg(alignment, c("l", "r", "ctr"), several.ok = TRUE )
  }

  header_str <- character(length = 0L)
  if( header ){
    header_str  <- paste0(
      sprintf("<a:tr h=\"%.0f\">", row_height),
      paste0(as_tc(colnames(x), align = alignment), collapse = ""),
      "</a:tr>"
    )
  }

  z <- mapply(as_tc, x, alignment, SIMPLIFY = FALSE)
  z <- do.call(paste0, z)
  z <- paste0(sprintf("<a:tr h=\"%.0f\">", row_height), z, "</a:tr>", collapse = "")

  z <- paste0(str, header_str, z, "</a:tbl>")
  z <- paste0(
    "<a:graphic>",
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">",
    z, "</a:graphicData>", "</a:graphic>")
}


#' @export
#' @title Table block
#' @description Create a representation of a table
#' @param x a data.frame to add as a table
#' @param header display header if TRUE
#' @param properties table properties, see [prop_table()].
#' Table properties are not handled identically between Word and PowerPoint
#' output format. They are fully supported with Word but for PowerPoint (which
#' does not handle as many things as Word for tables), only conditional
#' formatting properties are supported.
#' @param alignment alignment for each columns, 'l' for left, 'r' for right
#' and 'c' for center. Default to NULL.
#' @examples
#' block_table(x = head(iris))
#'
#' block_table(x = mtcars, header = TRUE,
#'   properties = prop_table(
#'     tcf = table_conditional_formatting(
#'       first_row = TRUE, first_column = TRUE)
#'   ))
#' @family block functions for reporting
#' @seealso [prop_table()]
block_table <- function(x, header = TRUE, properties = prop_table(), alignment = NULL) {

  stopifnot(is.data.frame(x))
  if(inherits(x, "tbl_df"))
    x <- as.data.frame(
      x, check.names = FALSE, stringsAsFactors = FALSE )

  z <- list(
    x = x,
    header = header,
    properties = properties,
    alignment = alignment
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

  value <- characterise_df(x$x)

  out <- table_docx(
    x = value, header = x$header,
    properties = x$properties,
    alignment = x$alignment,
    base_document = base_document,
    add_ns = add_ns
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
  value_str <- table_pptx(value, style_id = x$properties$style,
                          alignment = x$alignment,
                        col_width = as.integer((width/ncol(x$x))*914400),
                        row_height = as.integer((height/nrow(x$x))*914400),
                        tcf = x$properties$tcf,
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
#' @title Concatenate formatted text as a paragraph
#' @description Create a paragraph representation by concatenating
#' formatted text or images. The result can be inserted in a Word document
#' or a PowerPoint presentation and can also be inserted in a [block_list()]
#' call.
#'
#' All its arguments will be concatenated to create a paragraph where chunks of
#' text and images are associated with formatting properties.
#'
#' \code{fpar} supports [ftext()], [external_img()], \code{run_*} functions
#' (i.e. [run_autonum()], [run_seqfield()]) when output is Word, and simple strings.
#'
#' Default text and paragraph formatting properties can also be modified
#' with function `update()`.
#'
#' @param ... cot objects ([ftext()], [external_img()])
#' @param fp_p paragraph formatting properties, see [fp_par()]
#' @param fp_t default text formatting properties. This is used as
#' text formatting properties when simple text is provided as argument,
#' see [fp_text()].
#' @param values a list of cot objects. If provided, argument \code{...} will be ignored.
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
#' @seealso [block_list()], [body_add_fpar()], [ph_with()]
fpar <- function( ..., fp_p = fp_par(), fp_t = fp_text_lite(), values = NULL) {
  out <- list()

  if( is.null(values)){
    values <- list(...)
  }

  out$chunks <- values
  out$fp_p <- fp_p
  out$fp_t <- fp_t
  class(out) <- c("fpar", "block")
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
    "<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>")

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
#' @title List of blocks
#' @description A list of blocks can be used to gather
#' several blocks (paragraphs, tables, ...) into a single
#' object. The result can be added into a Word document or a
#' PowerPoint presentation.
#' @param ... a list of blocks. When output is only for
#' Word, objects of class [external_img()] can
#' also be used in fpar construction to mix text and images
#' in a single paragraph. Supported objects are:
#' [block_caption()], [block_pour_docx()], [block_section()],
#' [block_table()], [block_toc()], [fpar()], [plot_instr()].
#' @examples
#'
#' @example examples/block_list.R
#' @seealso [ph_with()], [body_add_blocks()], [fpar()]
#' @family block functions for reporting
block_list <- function(...){
  x <- list(...)
  z <- list()
  for(i in x){
    if(inherits(i, "block")) {
      z <- append(z, list(i))
    } else if(is.character(i)){
      z <- append(z, lapply(i, fpar))
    }
  }
  class(z) <- c("block_list", "block")
  z
}

#' @export
to_wml.block_list <- function(x, add_ns = FALSE, ...) {
  out <- character(length(x))
  for(i in seq_along(x) ){
    out[i] <- to_wml(x[[i]], add_ns = add_ns)
  }

  paste0(out, collapse = "")
}

#' @export
to_html.block_list <- function(x, add_ns = FALSE, ...) {
  str <- vapply(x, to_html, NA_character_)
  paste0(str, collapse = "")
}

# unordered list ----
#' @export
#' @title Unordered list
#' @description unordered list of text for PowerPoint
#' presentations. Each text is associated with
#' a hierarchy level.
#' @param str_list list of strings to be included in the object
#' @param level_list list of levels for hierarchy structure. Use
#' 0 for 'no bullet', 1 for level 1, 2 for level 2 and so on.
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
  tmpl <- "%s<a:pPr%s>%s</a:pPr><a:r>%s<a:t>%s</a:t></a:r></a:p>"
  lvl <- sprintf(" lvl=\"%.0f\"", x$lvl - 1)
  lvl <- ifelse(x$lvl > 1, lvl, "")
  bu_none <- ifelse(x$lvl < 1, "<a:buNone/>", "")
  p <- sprintf(tmpl, open_tag, lvl, bu_none, style_str, htmlEscapeCopy(x$str) )
  p <- paste(p, collapse = "")
  p
}

# plot_instr -----
#' @title Wrap plot instructions for png plotting in Powerpoint or Word
#' @description A simple wrapper to capture
#' plot instructions that will be executed and copied in a document. It produces
#' an object of class 'plot_instr' with a corresponding method [ph_with()] and
#' [body_add_plot()].
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
#' @seealso [ph_with()], [body_add_plot()]
#' @family block functions for reporting
plot_instr <- function(code) {
  out <- list()
  out$code <- substitute(code)
  class(out) <- "plot_instr"
  return(out)
}

