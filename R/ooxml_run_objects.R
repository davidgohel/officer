# generics -----

#' @export
#' @title Convert officer objects to WordprocessingML
#' @description Convert an object made with package officer
#' to WordprocessingML. The result is a string.
#' @param x object
#' @param add_ns should namespace be added to the top tag
#' @param ... Arguments to be passed to methods
#' @family functions for officer extensions
#' @keywords internal
to_wml <- function(x, add_ns = FALSE, ...) {
  UseMethod("to_wml")
}

#' @export
#' @title Convert officer objects to PresentationML
#' @description Convert an object made with package officer
#' to PresentationML. The result is a string.
#' @param x object
#' @param add_ns should namespace be added to the top tag
#' @param ... Arguments to be passed to methods
#' @family functions for officer extensions
#' @keywords internal
to_pml <- function(x, add_ns = FALSE, ...) {
  UseMethod("to_pml")
}

#' @export
#' @title Convert officer objects to HTML
#' @description Convert an object made with package officer
#' to HTML. The result is a string.
#' @param x object
#' @param ... Arguments to be passed to methods
#' @keywords internal
to_html <- function(x, ...) {
  UseMethod("to_html")
}


# bookmark -----

bookmark <- function(id, str) {
  new_id <- UUIDgenerate()
  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\"/>", new_id, id)
  bm_start_end <- sprintf("<w:bookmarkEnd w:id=\"%s\"/>", new_id)
  paste0(bm_start_str, str, bm_start_end)
}

# ftext -----
#' @export
#' @title formatted chunk of text
#' @description Format a chunk of text with text formatting properties (bold, color, ...).
#' The function allows you to create pieces of text formatted the way you want.
#' @param text text value, a single character value
#' @param prop formatting text properties returned by [fp_text]. It also can be NULL in
#' which case, no formatting is defined (the default is applied).
#' @section usage:
#' You can use this function in conjunction with [fpar] to create paragraphs
#' consisting of differently formatted text parts. You can also use this
#' function as an *r chunk* in an R Markdown document made with package
#' officedown.
#' @examples
#' ftext("hello", fp_text())
#'
#' properties1 <- fp_text(color = "red")
#' properties2 <- fp_text(bold = TRUE, shading.color = "yellow")
#' ftext1 <- ftext("hello", properties1)
#' ftext2 <- ftext("World", properties2)
#' paragraph <- fpar(ftext1, " ", ftext2)
#'
#' x <- read_docx()
#' x <- body_add(x, paragraph)
#' print(x, target = tempfile(fileext = ".docx"))
#' @seealso [fp_text]
#' @family run functions for reporting
ftext <- function(text, prop = NULL) {
  if(is.character(text)) {
    value <- enc2utf8(text)
  } else {
    value <- formatC(text)
  }

  out <- list( value = value, pr = prop )
  class(out) <- c("ftext", "cot", "run")
  out
}

#' @export
print.ftext <- function (x, ...) {
  cat("text: ", x$value, "\nformat:\n", sep = "")
  print(x$pr)
}

#' @export
to_wml.ftext <- function(x, add_ns = FALSE, ...) {
  out <- paste0("<w:r>", if(!is.null(x$pr)) rpr_wml(x$pr),
                "<w:t xml:space=\"preserve\">",
                htmlEscapeCopy(x$value), "</w:t></w:r>")
  out
}

#' @export
to_pml.ftext <- function(x, add_ns = FALSE, ...) {
  open_tag <- ar_ns_no
  if (add_ns) {
    open_tag <- ar_ns_yes
  }
  paste0(open_tag, if(!is.null(x$pr)) rpr_pml(x$pr),
         "<a:t>", htmlEscapeCopy(x$value), "</a:t></a:r>")
}

#' @export
to_html.ftext <- function(x, ...) {
  sprintf("<span style=\"%s\">%s</span>", if(!is.null(x$pr)) rpr_css(x$pr) else "", htmlEscapeCopy(x$value))
}


# Word computed field ----
#' @export
#' @title seqfield
#' @description Create a Word computed field.
#' @note
#' In the previous version, this function was called `run_seqfield`
#' but the name was wrong and should have been `run_word_field`.
#' @inheritSection ftext usage
#' @param field,seqfield computed field string (`seqfield` will be
#' totally superseded by `field` in the futur).
#' @param prop formatting text properties returned by [fp_text].
#' @examples
#' run_word_field(field = "PAGE  \\* MERGEFORMAT")
#' run_word_field(field = "Date \\@ \"MMMM d yyyy\"")
#' @family run functions for reporting
#' @family Word computed fields
run_word_field <- function(field, prop = NULL, seqfield = field) {
  z <- list(
    seqfield = seqfield, pr = prop
  )
    class(z) <- c("run_seqfield", "run")
  z
}

#' @export
#' @rdname run_word_field
run_seqfield <- run_word_field

#' @export
to_wml.run_seqfield <- function(x, add_ns = FALSE, ...) {

  pr <- if(!is.null(x$pr)) rpr_wml(x$pr) else "<w:rPr/>"

  xml_elt_1 <- paste0(
    wr_ns_yes,
    pr,
    "<w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/>",
    "</w:r>"
  )
  xml_elt_2 <- paste0(
    wr_ns_yes,
    pr,
    sprintf("<w:instrText xml:space=\"preserve\" w:dirty=\"true\">%s</w:instrText>", x$seqfield),
    "</w:r>"
  )
  xml_elt_3 <- paste0(
    wr_ns_yes,
    pr,
    "<w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/>",
    "</w:r>"
  )
  out <- paste0(xml_elt_1, xml_elt_2, xml_elt_3)

  out
}

# autonum ----

#' @export
#' @title auto number
#' @description Create an autonumbered chunk, i.e. a string
#' representation of a sequence, each item will be numbered.
#' These runs can also be bookmarked and be used later for
#' cross references.
#' @inheritSection ftext usage
#' @param seq_id sequence identifier
#' @param pre_label,post_label text to add before and after number
#' @param bkm bookmark id to associate with autonumber run. If NULL, no bookmark
#' is added. Value can only be made of alpha numeric characters, '-' and '_'.
#' @param bkm_all if TRUE, the bookmark will be set on the whole string, if
#' FALSE, the bookmark will be set on the number only. Default to FALSE.
#' As an effect when a reference to this bookmark is used, the text can
#' be like "Table 1" or "1" (pre_label is not included in the referenced
#' text).
#' @param prop formatting text properties returned by [fp_text].
#' @examples
#' run_autonum()
#' run_autonum(seq_id = "fig", pre_label = "fig. ")
#' run_autonum(seq_id = "tab", pre_label = "Table ", bkm = "anytable")
#' @family run functions for reporting
#' @family Word computed fields
run_autonum <- function(seq_id = "table", pre_label = "Table ", post_label = ": ",
                        bkm = NULL, bkm_all = FALSE, prop = NULL) {
  bkm <- check_bookmark_id(bkm)
  z <- list(
    seq_id = seq_id,
    pre_label = pre_label,
    post_label = post_label,
    bookmark = bkm,
    bookmark_all = bkm_all,
    pr = prop
  )
  class(z) <- c("run_autonum", "run")

  z
}

#' @export
to_wml.run_autonum <- function(x, add_ns = FALSE, ...) {
  if(is.null(x$seq_id)) return("")

  pr <- if(!is.null(x$pr)) rpr_wml(x$pr) else "<w:rPr/>"

  run_str_pre <- sprintf("<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>", pr, x$pre_label)
  run_str_post <- sprintf("<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>", pr, x$post_label)
  sqf <- run_seqfield(seqfield = paste0("SEQ ", x$seq_id, " \u005C* Arabic"), prop = x$pr)
  sf_str <- to_wml(sqf)
  if(!is.null(x$bookmark)){
    if(x$bookmark_all){
      out <- paste0(
        bookmark(id = x$bookmark,
                 str = paste0(run_str_pre, sf_str)),
        run_str_post)
    } else {
      sf_str <- bookmark(id = x$bookmark, sf_str)
      out <- paste0(run_str_pre, sf_str, run_str_post)
    }
  } else {
    out <- paste0(run_str_pre, sf_str, run_str_post)
  }

  out
}

# reference ----

#' @export
#' @title reference
#' @description Create a representation of a reference
#' @param id reference id, a string
#' @param prop formatting text properties returned by [fp_text].
#' @inheritSection ftext usage
#' @examples
#' run_reference('a_ref')
#' @family run functions for reporting
#' @family Word computed fields
run_reference <- function(id, prop = NULL) {
  z <- paste0(" REF ", id, " \\h ")

  z <- list(
    id = id, pr = prop
  )
  class(z) <- c("run_reference", "run")

  z
}

#' @export
to_wml.run_reference <- function(x, add_ns = FALSE, ...) {
  out <- to_wml(run_seqfield(seqfield = paste0(" REF ", x$id, " \\h "), prop = x$pr))
  out
}

# hyperlink text ----

#' @export
#' @title formatted chunk of text with hyperlink
#' @description Format a chunk of text with text formatting properties (bold, color, ...),
#' the chunk is associated with an hyperlink.
#' @param text text value, a single character value
#' @param prop formatting text properties returned by [fp_text]. It also can be NULL in
#' which case, no formatting is defined (the default is applied).
#' @param href URL value
#' @inheritSection ftext usage
#' @examples
#' ft <- fp_text(font.size = 12, bold = TRUE)
#' hyperlink_ftext(
#'   href = "https://cran.r-project.org/index.html",
#'   text = "some text", prop = ft)
#' @family run functions for reporting
hyperlink_ftext <- function(text, prop = NULL, href) {
  if(is.character(text)) {
    value <- enc2utf8(text)
  } else {
    value <- formatC(text)
  }

  out <- list(value = value, pr = prop, href = href)
  class(out) <- c("hyperlink_ftext", "cot", "run")
  out
}

#' @export
to_wml.hyperlink_ftext <- function(x, add_ns = FALSE, ...) {

  out <- paste0("<w:r>", if(!is.null(x$pr)) rpr_wml(x$pr),
                "<w:t xml:space=\"preserve\">",
                htmlEscapeCopy(x$value), "</w:t></w:r>")
  paste0("<w:hyperlink r:id=\"", htmlEscapeCopy(x$href), "\">", out, "</w:hyperlink>")
}

#' @export
to_pml.hyperlink_ftext <- function(x, add_ns = FALSE, ...) {
  open_tag <- ar_ns_no
  if (add_ns) {
    open_tag <- ar_ns_yes
  }
  str_ <- sprintf("<a:hlinkClick r:id=\"%s\"/>", x$href)

  if(!is.null(x$pr)) {
    rpr_pml_ <- rpr_pml(x$pr)
    m <- regexpr(">", rpr_pml_)
    regmatches(rpr_pml_, m) <- paste0(">",str_)
  } else {
    rpr_pml_ <- paste0("<a:rPr>", str_, "</a:rPr>")
  }

  paste0(open_tag, rpr_pml_,
         "<a:t>", htmlEscapeCopy(x$value), "</a:t></a:r>")
}

#' @export
to_html.hyperlink_ftext <- function(x, ...) {
  out <- sprintf("<span style=\"%s\">%s</span>",
                 if(!is.null(x$pr)) rpr_css(x$pr) else "",
                 htmlEscapeCopy(x$value))
  sprintf("<a href=\"%s\">%s</a>", x$href, out)
  out
}



# bookmark ----

#' @export
#' @title bookmark for Word
#' @description Add a bookmark on a run object.
#' @param bkm bookmark id to associate with run. Value can only be
#' made of alpha numeric characters, '-' and '_'.
#' @param run a run object, made with a call to one of
#' the "run functions for reporting".
#' @inheritSection ftext usage
#' @examples
#' ft <- fp_text(font.size = 12, bold = TRUE)
#' run_bookmark("par1", ftext("some text", ft))
#' @family run functions for reporting
run_bookmark <- function(bkm, run) {

  all_run <- FALSE

  if(inherits(run, "run")){
    run <- list(run)
  }

  if(is.list(run) && !inherits(run, "run")){
    all_run <- all(vapply(run, inherits, FUN.VALUE = FALSE, what = "run" ))
  }


  if(!all_run)
    stop("`run` must be a run object (ftext for example) or a list of run objects.")

  bkm <- check_bookmark_id(bkm)

  z <- list(id = bkm, run = run)
  class(z) <- c("run_bookmark", "run")
  z
}

#' @export
to_wml.run_bookmark <- function(x, add_ns = FALSE, ...) {
  runs <- lapply(x$run, to_wml, add_ns = add_ns, ...)
  runs <- do.call(paste0, runs)
  bookmark(id = x$id, str = runs)
}

# breaks ----

#' @export
#' @title page break for Word
#' @description Object representing a page break for a Word document.
#' @inheritSection ftext usage
#' @examples
#'
#' @example examples/run_pagebreak.R
#' @family run functions for reporting
run_pagebreak <- function() {
  z <- list()
  class(z) <- c("run_pagebreak", "run")
  z
}

#' @export
to_wml.run_pagebreak <- function(x, add_ns = FALSE, ...) {
  out <- "<w:r><w:br w:type=\"page\"/></w:r>"
  out
}

#' @export
#' @title column break
#' @description Create a representation of a column break
#' @inheritSection ftext usage
#' @examples
#' run_columnbreak()
#' @family run functions for reporting
run_columnbreak <- function() {
  z <- list()
  class(z) <- c("run_columnbreak", "run")

  z
}

#' @export
to_wml.run_columnbreak <- function(x, add_ns = FALSE, ...) {
  out <- "<w:r><w:br w:type=\"column\"/></w:r>"
  out
}

#' @export
#' @title page break for Word
#' @description Object representing a line break for a Word document. The
#' result must be used within a call to [fpar].
#' @inheritSection ftext usage
#' @examples
#'
#' @example examples/run_linebreak.R
#' @family run functions for reporting
run_linebreak <- function() {
  z <- list()
  class(z) <- c("run_linebreak", "run")

  z
}

#' @export
to_wml.run_linebreak <- function(x, add_ns = FALSE, ...) {
  out <- "<w:r><w:br w:type=\"textWrapping\"/></w:r>"
  out
}

# sections ----
inch_to_tweep <- function(x) round(x * 72 * 20, digits = 0)

#' @export
#' @title page size object
#' @description The function creates a representation of the dimensions of a page. The
#' dimensions are defined by length, width and orientation. If the orientation is
#' in landscape mode then the length becomes the width and the width becomes the length.
#' @param width,height page width, page height (in inches).
#' @param orient page orientation, either 'landscape', either 'portrait'.
#' @examples
#' page_size(orient = "landscape")
#' @family functions for section definition
page_size <- function(width = 21 / 2.54, height = 29.7 / 2.54, orient = "portrait") {
  if (orient %in% "portrait") {
    h <- max(c(height, width))
    w <- min(c(height, width))
  } else {
    h <- min(c(height, width))
    w <- max(c(height, width))
  }

  z <- list(width = w, height = h, orient = orient)
  class(z) <- c("page_size")
  z
}

#' @export
to_wml.page_size <- function(x, add_ns = FALSE, ...) {
  out <- sprintf(
    "<w:pgSz w:h=\"%.0f\" w:w=\"%.0f\" w:orient=\"%s\"/>",
    inch_to_tweep(x$height), inch_to_tweep(x$width), x$orient
  )
  out
}

#' @export
#' @title section columns
#' @description The function creates a representation of the columns of a section.
#' @param widths columns widths in inches. If 3 values, 3 columns
#' will be produced.
#' @param space space in inches between columns.
#' @param sep if TRUE a line is separating columns.
#' @examples
#' section_columns()
#' @family functions for section definition
section_columns <- function(widths = c(2.5,2.5), space = .25, sep = FALSE) {
  if( length(widths) < 2 )
    stop("length of widths should be at least 2")

  z <- list(widths = widths, space = space, sep = sep)
  class(z) <- c("section_columns")
  z
}
#' @export
to_wml.section_columns <- function(x, add_ns = FALSE, ...) {
  widths <- x$widths * 20 * 72
  space <- x$space * 20 * 72

  columns_str_all_but_last <- sprintf("<w:col w:w=\"%.0f\" w:space=\"%.0f\"/>",
                                      widths[-length(widths)], space)
  columns_str_last <- sprintf("<w:col w:w=\"%.0f\"/>",
                              widths[length(widths)])
  columns_str <- c(columns_str_all_but_last, columns_str_last)

  sprintf("<w:cols w:num=\"%.0f\" w:sep=\"%.0f\" w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
                         length(widths), as.integer(x$sep),
                         space, paste0(columns_str, collapse = "") )
}

#' @export
#' @title page margins object
#' @description The margins for each page of a sectionThe function creates a representation of the dimensions of a page. The
#' dimensions are defined by length, width and orientation. If the orientation is
#' in landscape mode then the length becomes the width and the width becomes the length.
#' @param bottom,top distance (in inches) between the bottom/top of the text margin and the bottom/top of the page. The text is placed at the greater of the value of this attribute and the extent of the header/footer text. A negative value indicates that the content should be measured from the bottom/topp of the page regardless of the footer/header, and so will overlap the footer/header. For example, `header=-0.5, bottom=1` means that the footer must start one inch from the bottom of the page and the main document text must start a half inch from the bottom of the page. In this case, the text and footer overlap since bottom is negative.
#' @param footer distance (in inches) from the bottom edge of the page to the bottom edge of the footer.
#' @param header distance (in inches) from the top edge of the page to the top edge of the header.
#' @param left,right distance (in inches) from the left/right edge of the page to the left/right edge of the text.
#' @param gutter page gutter (in inches).
#' @examples
#' page_mar()
#' @family functions for section definition
page_mar <- function(bottom = 1, top = 1, right = 1, left = 1,
                     header = 0.5, footer = 0.5, gutter = 0.5) {
  z <- list(header = header, bottom = bottom, top = top, right = right, left = left, footer = footer, gutter = gutter)
  class(z) <- c("page_mar")
  z
}
#' @export
to_wml.page_mar <- function(x, add_ns = FALSE, ...) {
  out <- sprintf(
    "<w:pgMar w:header=\"%.0f\" w:bottom=\"%.0f\" w:top=\"%.0f\" w:right=\"%.0f\" w:left=\"%.0f\" w:footer=\"%.0f\" w:gutter=\"%.0f\"/>",
    inch_to_tweep(x$header), inch_to_tweep(x$bottom), inch_to_tweep(x$top),
    inch_to_tweep(x$right), inch_to_tweep(x$left), inch_to_tweep(x$footer),
    inch_to_tweep(x$gutter)
  )

  out
}

docx_section_type <- c("continuous", "evenPage", "nextColumn", "nextPage", "oddPage")

#' @export
#' @title section properties
#' @description A section is a grouping of blocks (ie. paragraphs and tables)
#' that have a set of properties that define pages on which the text will appear.
#'
#' A Section properties object stores information about page composition,
#' such as page size, page orientation, borders and margins.
#' @note
#' There is no support yet for header and footer contents definition.
#' @param page_size page dimensions, an object generated with function [page_size].
#' @param page_margins page margins, an object generated with function [page_mar].
#' @param type Section type. It defines how the contents of the section will be
#' placed relative to the previous section. Available types are "continuous"
#' (begins the section on the next paragraph), "evenPage" (begins on the next
#' even-numbered page), "nextColumn" (begins on the next column on the page),
#' "nextPage" (begins on the following page), "oddPage" (begins on the next
#' odd-numbered page).
#' @param section_columns section columns, an object generated with function [section_columns].
#' @examples
#' library(officer)
#'
#' landscape_one_column <- block_section(
#'   prop_section(
#'     page_size = page_size(orient = "landscape"), type = "continuous"
#'   )
#' )
#' landscape_two_columns <- block_section(
#'   prop_section(
#'     page_size = page_size(orient = "landscape"), type = "continuous",
#'     section_columns = section_columns(widths = c(4.75, 4.75))
#'   )
#' )
#'
#' doc_1 <- read_docx()
#' # there starts section with landscape_one_column
#' doc_1 <- body_add_table(doc_1, value = mtcars[1:10,], style = "table_template")
#' doc_1 <- body_end_block_section(doc_1, value = landscape_one_column)
#' # there stops section with landscape_one_column
#'
#'
#' # there starts section with landscape_two_columns
#' doc_1 <- body_add_par(doc_1, value = paste(rep(letters, 50), collapse = " "))
#' doc_1 <- body_end_block_section(doc_1, value = landscape_two_columns)
#' # there stops section with landscape_two_columns
#'
#' doc_1 <- body_add_table(doc_1, value = mtcars[1:25,], style = "table_template")
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#' @seealso [block_section]
#' @family functions for section definition
#' @section Illustrations:
#'
#' \if{html}{\figure{prop_section_doc_1.png}{options: width=80\%}}
prop_section <- function(page_size = NULL, page_margins = NULL,
                         type = NULL, section_columns = NULL) {
  z <- list()
  if (!is.null(type) && !type %in% docx_section_type) {
    stop("type must be one of ", paste(shQuote(docx_section_type), collapse = ", "))
  }
  z$type <- type

  if (!is.null(page_size) && !inherits(page_size, "page_size")) {
    stop("page_size must be a page_size object")
  }
  z$page_size <- page_size

  if (!is.null(page_margins) && !inherits(page_margins, "page_mar")) {
    stop("page_margins must be a page_mar object")
  }
  z$page_margins <- page_margins

  if (!is.null(section_columns) && !inherits(section_columns, "section_columns")) {
    stop("section_columns must be a section_columns object")
  }
  z$section_columns <- section_columns

  z <- Filter(function(x) !is.null(x), z)
  class(z) <- c("prop_section", "prop")

  z
}

#' @export
to_wml.prop_section <- function(x, add_ns = FALSE, ...) {

  paste0(
    "<w:sectPr",
    if (add_ns) " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
    ">",
    if(!is.null(x$page_margins)) to_wml(x$page_margins),
    if(!is.null(x$page_size)) to_wml(x$page_size),
    if(!is.null(x$type)) "<w:type w:val=\"", x$type, "\"/>",
    if(!is.null(x$section_columns)) to_wml(x$section_columns) else "<w:cols/>",
    "</w:sectPr>"
  )
}


# external_img ----
#' @export
#' @title external image
#' @description Wraps an image in an object that can then be embedded
#' in a PowerPoint slide or within a Word paragraph.
#'
#' The image is added as a shape in PowerPoint (it is not possible to mix text and
#' images in a PowerPoint form). With a Word document, the image will be
#' added inside a paragraph.
#' @param src image file path
#' @param width height in inches.
#' @param height height in inches
#' @param alt alternative text for images
#' @inheritSection ftext usage
#' @examples
#'
#' @example examples/external_img.R
#' @seealso [ph_with], [body_add], [fpar]
#' @family run functions for reporting
external_img <- function(src, width = .5, height = .2, alt = "") {
  # note: should it be vectorized
  check_src <- all(grepl("^rId", src)) || all(file.exists(src))
  if( !check_src ){
    stop("src must be a string starting with 'rId' or an image filename")
  }

  class(src) <- c("external_img", "cot", "run")
  attr(src, "dims") <- list(width = width, height = height)
  attr(src, "alt") <- alt
  src
}

dim.external_img <- function( x ){
  x <- attr(x, "dims")
  data.frame(width = x$width, height = x$height)
}

as.data.frame.external_img <- function( x, ... ){
  dimx <- attr(x, "dims")
  data.frame(path = as.character(x), width = dimx$width, height = dimx$height, alt = attr(x, "alt"), stringsAsFactors = FALSE)
}


pic_pml <- function( left = 0, top = 0, width = 3, height = 3,
                     bg = "transparent", rot = 0, label = "", ph = "<p:ph/>", src, alt_text = ""){

  if( !is.null(bg) && !is.color( bg ) )
    stop("bg must be a valid color.", call. = FALSE )

  bg_str <- solid_fill_pml(bg)

  if (missing(alt_text)) alt_text <- ""


  xfrm_str <- a_xfrm_str(left = left, top = top, width = width, height = height, rot = rot)
  if( is.null(ph) || is.na(ph)){
    ph = "<p:ph/>"
  }
  blipfill <- paste0(
    "<p:blipFill>",
    sprintf("<a:blip cstate=\"print\" r:embed=\"%s\"/>", src),
    "<a:stretch><a:fillRect/></a:stretch>",
    "</p:blipFill>")
  str <- "
<p:pic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">
  <p:nvPicPr>
    <p:cNvPr id=\"0\" name=\"%s\" descr=\"%s\"/>
    <p:cNvPicPr><a:picLocks noGrp=\"1\"/></p:cNvPicPr>
    <p:nvPr>%s</p:nvPr>
  </p:nvPicPr>
  %s
  <p:spPr>%s<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>%s</p:spPr>
</p:pic>
"
  sprintf(str, label, alt_text, ph, blipfill, xfrm_str, bg_str )

}


#' @export
to_wml.external_img <- function(x, add_ns = FALSE, ...) {

  open_tag <- wr_ns_no
  if (add_ns) {
    open_tag <- wr_ns_yes
  }

  dims <- attr(x, "dims")
  width <- dims$width
  height <- dims$height


  paste0(open_tag,
         "<w:rPr/><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">",
         sprintf("<wp:extent cx=\"%.0f\" cy=\"%.0f\"/>", width * 12700*72, height * 12700*72),
         "<wp:docPr id=\"\" name=\"\" descr=\"", attr(x, "alt"),"\"/>",
         "<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>",
         "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
         "<pic:nvPicPr>",
         "<pic:cNvPr id=\"\" name=\"\"/>",
         "<pic:cNvPicPr><a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>",
         "</pic:cNvPicPr></pic:nvPicPr>",
         "<pic:blipFill>",
         sprintf("<a:blip r:embed=\"%s\"/>", as.character(x)),
         "<a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>",
         "<pic:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/>",
         sprintf("<a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></pic:spPr>", width * 12700, height * 12700),
         "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
  )

}

#' @export
to_html.external_img <- function(x, ...) {
  dims <- attr(x, "dims")
  width <- dims$width
  height <- dims$height

  if(!requireNamespace("base64enc")) {
    stop("package 'base64enc' is required for this function, please install it.")
  }

  if( grepl("\\.png", ignore.case = TRUE, x = x) ){
    mime <- "image/png"
  } else if( grepl("\\.gif", ignore.case = TRUE, x = x) ){
    mime <- "image/gif"
  } else if( grepl("\\.jpg", ignore.case = TRUE, x = x) ){
    mime <- "image/jpeg"
  } else if( grepl("\\.jpeg", ignore.case = TRUE, x = x) ){
    mime <- "image/jpeg"
  } else if( grepl("\\.svg", ignore.case = TRUE, x = x) ){
    mime <- "image/svg+xml"
  } else if( grepl("\\.tiff", ignore.case = TRUE, x = x) ){
    mime <- "image/tiff"
  } else if( grepl("\\.webp", ignore.case = TRUE, x = x) ){
    mime <- "image/webp"
  } else {
    stop("this format is not implemented")
  }
  x <- base64enc::dataURI(file = as.character(x), mime = mime )
  sprintf("<img style=\"vertical-align:middle;width:%.0fpx;height:%.0fpx;\" src=\"%s\" alt=\"%s\"/>", width*72, height*72, x, attr(x, "alt"))

}


