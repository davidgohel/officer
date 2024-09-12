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
#' @family functions for officer extensions
#' @keywords internal
to_html <- function(x, ...) {
  UseMethod("to_html")
}


# bookmark -----

bookmark <- function(id, str) {
  new_id <- uuid_generate()
  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\"/>", new_id, id)
  bm_start_end <- sprintf("<w:bookmarkEnd w:id=\"%s\"/>", new_id)
  paste0(bm_start_str, str, bm_start_end)
}

# ftext -----
#' @export
#' @title Formatted chunk of text
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
  if (is.character(text)) {
    value <- enc2utf8(text)
  } else {
    value <- formatC(text)
  }

  out <- list(value = value, pr = prop)
  class(out) <- c("ftext", "cot", "run")
  out
}

#' @export
print.ftext <- function(x, ...) {
  cat("text: ", x$value, "\nformat:\n", sep = "")
  print(x$pr)
}

#' @export
to_wml.ftext <- function(x, add_ns = FALSE, ...) {
  out <- paste0(
    "<w:r>", if (!is.null(x$pr)) rpr_wml(x$pr),
    "<w:t xml:space=\"preserve\">",
    htmlEscapeCopy(x$value), "</w:t></w:r>"
  )
  out
}

#' @export
to_pml.ftext <- function(x, add_ns = FALSE, ...) {
  open_tag <- ar_ns_no
  if (add_ns) {
    open_tag <- ar_ns_yes
  }
  paste0(
    open_tag, if (!is.null(x$pr)) rpr_pml(x$pr),
    "<a:t>", htmlEscapeCopy(x$value), "</a:t></a:r>"
  )
}

#' @export
to_html.ftext <- function(x, ...) {
  sprintf("<span style=\"%s\">%s</span>", if (!is.null(x$pr)) rpr_css(x$pr) else "", htmlEscapeCopy(x$value))
}


# Word chunk ----
#' @export
#' @title Word chunk of text with a style
#' @description Format a chunk of text associated with a 'Word' character style.
#' The style is defined with its unique identifer.
#' @param text text value, a single character value
#' @param style_id 'Word' unique style identifier associated with the style to use.
#' @examples
#' run1 <- run_wordtext("hello", "DefaultParagraphFont")
#' paragraph <- fpar(run1)
#'
#' x <- read_docx()
#' x <- body_add_fpar(x, paragraph)
#' print(x, target = tempfile(fileext = ".docx"))
#' @seealso [ftext()]
#' @family run functions for reporting
run_wordtext <- function(text, style_id = NULL) {
  if (is.character(text)) {
    value <- enc2utf8(text)
  } else {
    value <- formatC(text)
  }

  out <- list(value = value, style_id = style_id)
  class(out) <- c("run_wordtext", "cot", "run")
  out
}

#' @export
print.run_wordtext <- function(x, ...) {
  cat("text: `", x$value, "`\nword style id: `", x$style_id, "`\n", sep = "")
}
#
#' @export
to_wml.run_wordtext <- function(x, add_ns = FALSE, ...) {
  out <- paste0(
    "<w:r>",
    "<w:rPr>",
    sprintf("<w:rStyle w:val=\"%s\"/>", x$style_id),
    "</w:rPr>",
    "<w:t xml:space=\"preserve\">",
    htmlEscapeCopy(x$value), "</w:t></w:r>"
  )
  out
}


# Word computed field ----
#' @export
#' @title 'Word' computed field
#' @description Create a 'Word' computed field.
#' @note
#' In the previous version, this function was called `run_seqfield`
#' but the name was wrong and should have been `run_word_field`.
#' @inheritSection ftext usage
#' @param field Value for a "Word Computed Field" as a string.
#' @param seqfield deprecated in favor of `field`.
#' @param prop formatting text properties returned by [fp_text].
#' @examples
#' run_word_field(field = "PAGE  \\* MERGEFORMAT")
#' run_word_field(field = "Date \\@ \"MMMM d yyyy\"")
#' @family run functions for reporting
#' @family Word computed fields
run_word_field <- function(field, prop = NULL, seqfield = NULL) {
  if (!is.null(seqfield)) {
    field <- seqfield
    message("`seqfield` argument is deprecated in favor of `field`")
  }
  z <- list(
    field = field, pr = prop
  )
  class(z) <- c("run_word_field", "run")
  z
}

#' @export
#' @rdname run_word_field
run_seqfield <- run_word_field

#' @export
to_wml.run_word_field <- function(x, add_ns = FALSE, ...) {
  pr <- if (!is.null(x$pr)) rpr_wml(x$pr) else "<w:rPr/>"

  xml_elt_1 <- paste0(
    wr_ns_yes,
    pr,
    "<w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/>",
    "</w:r>"
  )
  xml_elt_2 <- paste0(
    wr_ns_yes,
    pr,
    sprintf("<w:instrText xml:space=\"preserve\" w:dirty=\"true\">%s</w:instrText>", x$field),
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
#' @title Auto number
#' @description Create an autonumbered chunk, i.e. a string
#' representation of a sequence, each item will be numbered.
#' These runs can also be bookmarked and be used later for
#' cross references.
#' @inheritSection ftext usage
#' @param seq_id sequence identifier
#' @param pre_label,post_label text to add before and after number
#' @param bkm bookmark id to associate with autonumber run. If NULL, no bookmark
#' is added. Value can only be made of alpha numeric characters, ':', -' and '_'.
#' @param bkm_all if TRUE, the bookmark will be set on the whole string, if
#' FALSE, the bookmark will be set on the number only. Default to FALSE.
#' As an effect when a reference to this bookmark is used, the text can
#' be like "Table 1" or "1" (pre_label is not included in the referenced
#' text).
#' @param prop formatting text properties returned by [fp_text].
#' @param start_at If not NULL, it must be a positive integer, it
#' specifies the new number to use, at which number the auto
#' numbering will restart.
#' @param tnd *title number depth*, a positive integer (only applies if positive)
#' that specify the depth (or heading of level *depth*) to use for prefixing
#' the caption number with this last reference number. For example, setting
#' `tnd=2` will generate numbered captions like '4.3-2' (figure 2 of
#' chapter 4.3).
#' @param tns separator to use between title number and table
#' number. Default is "-".
#' @examples
#' run_autonum()
#' run_autonum(seq_id = "fig", pre_label = "fig. ")
#' run_autonum(seq_id = "tab", pre_label = "Table ", bkm = "anytable")
#' run_autonum(
#'   seq_id = "tab", pre_label = "Table ", bkm = "anytable",
#'   tnd = 2, tns = " "
#' )
#' @family run functions for reporting
#' @family Word computed fields
run_autonum <- function(seq_id = "table", pre_label = "Table ", post_label = ": ",
                        bkm = NULL, bkm_all = FALSE, prop = NULL,
                        start_at = NULL,
                        tnd = 0, tns = "-") {
  bkm <- check_bookmark_id(bkm)

  stopifnot(
    !is.null(tnd),
    !is.null(tns),
    is.numeric(tnd),
    sign(tnd) > -1,
    is.character(tns),
    is.null(start_at) || is.numeric(start_at),
    inherits(prop, "fp_text") || is.null(prop)
  )
  z <- list(
    seq_id = seq_id,
    pre_label = pre_label,
    post_label = post_label,
    bookmark = bkm,
    bookmark_all = bkm_all,
    pr = prop,
    start_at = start_at,
    tnd = tnd, tns = tns
  )
  class(z) <- c("run_autonum", "run")

  z
}

#' @export
#' @title Update bookmark of an autonumber run
#' @description This function lets recycling a object
#' made by [run_autonum()] by changing the bookmark value. This
#' is useful to avoid calling `run_autonum()` several times
#' because of many tables.
#' @param x an object of class [run_autonum()]
#' @param bkm bookmark id to associate with autonumber run.
#' Value can only be made of alpha numeric characters, ':', -' and '_'.
#' @examples
#' z <- run_autonum(
#'   seq_id = "tab", pre_label = "Table ",
#'   bkm = "anytable"
#' )
#' set_autonum_bookmark(z, bkm = "anothertable")
#' @seealso [run_autonum()]
set_autonum_bookmark <- function(x, bkm = NULL) {
  stopifnot(inherits(x, "run_autonum"))
  x$bookmark <- bkm
  x
}

#' @export
to_wml.run_autonum <- function(x, add_ns = FALSE, ...) {
  if (is.null(x$seq_id)) {
    return("")
  }

  pr <- if (!is.null(x$pr)) rpr_wml(x$pr) else "<w:rPr/>"

  run_str_pre <- sprintf("<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>", pr, x$pre_label)
  run_str_post <- sprintf("<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>", pr, x$post_label)

  seq_str <- paste0("SEQ ", x$seq_id, " \u005C* Arabic")
  if (!is.null(x$start_at) && is.numeric(x$start_at)) {
    seq_str <- paste(seq_str, "\\r", as.integer(x$start_at))
  }

  sqf <- run_word_field(field = seq_str, prop = x$pr)
  sf_str <- to_wml(sqf)

  if (x$tnd > 0) {
    z <- paste0(
      to_wml(run_word_field(field = paste0("STYLEREF ", x$tnd, " \\r"), prop = x$pr)),
      to_wml(ftext(x$tns, prop = x$pr))
    )
    sf_str <- paste0(z, sf_str)
  }


  if (!is.null(x$bookmark)) {
    if (x$bookmark_all) {
      out <- paste0(
        bookmark(
          id = x$bookmark,
          str = paste0(run_str_pre, sf_str)
        ),
        run_str_post
      )
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
#' @title Cross reference
#' @description Create a representation of a reference
#' @param id reference id, a string
#' @param prop formatting text properties returned by [fp_text].
#' @inheritSection ftext usage
#' @examples
#' run_reference("a_ref")
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
  out <- to_wml(run_word_field(field = paste0(" REF ", x$id, " \\h "), prop = x$pr))
  out
}

# hyperlink text ----

#' @export
#' @title Formatted chunk of text with hyperlink
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
#'   text = "some text", prop = ft
#' )
#' @family run functions for reporting
hyperlink_ftext <- function(text, prop = NULL, href) {
  if (is.character(text)) {
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
  out <- paste0(
    "<w:r>", if (!is.null(x$pr)) rpr_wml(x$pr),
    "<w:t xml:space=\"preserve\">",
    x$value, "</w:t></w:r>"
  )
  paste0("<w:hyperlink r:id=\"", officer_url_encode(x$href), "\">", out, "</w:hyperlink>")
}

#' @export
to_pml.hyperlink_ftext <- function(x, add_ns = FALSE, ...) {
  open_tag <- ar_ns_no
  if (add_ns) {
    open_tag <- ar_ns_yes
  }
  str_ <- sprintf("<a:hlinkClick r:id=\"%s\"/>", officer_url_encode(x$href))

  if (!is.null(x$pr)) {
    rpr_pml_ <- rpr_pml(x$pr)
    m <- regexpr(">", rpr_pml_)
    regmatches(rpr_pml_, m) <- paste0(">", str_)
  } else {
    rpr_pml_ <- paste0("<a:rPr>", str_, "</a:rPr>")
  }

  paste0(
    open_tag, rpr_pml_,
    "<a:t>", htmlEscapeCopy(x$value), "</a:t></a:r>"
  )
}

#' @export
to_html.hyperlink_ftext <- function(x, ...) {
  if (!is.null(x$pr)) {
    rpr_css <- paste0(" style=\"", rpr_css(x$pr), "\"")
  } else {
    rpr_css <- ""
  }

  sprintf("<a href=\"%s\"><span%s></span></a>", x$href, rpr_css)
}


# bookmark ----

#' @export
#' @title Bookmark for 'Word'
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

  if (inherits(run, "run")) {
    run <- list(run)
  }

  if (is.list(run) && !inherits(run, "run")) {
    all_run <- all(vapply(run, inherits, FUN.VALUE = FALSE, what = "run"))
  }


  if (!all_run) {
    stop("`run` must be a run object (ftext for example) or a list of run objects.")
  }

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
#' @title Page break for 'Word'
#' @description Object representing a page break for a Word document.
#' @inheritSection ftext usage
#' @examples
#' fp_t <- fp_text(font.size = 12, bold = TRUE)
#' an_fpar <- fpar("let's add a break page", run_pagebreak(), ftext("and blah blah!", fp_t))
#'
#' x <- read_docx()
#' x <- body_add(x, an_fpar)
#' print(x, target = tempfile(fileext = ".docx"))
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
#' @title Column break for 'Word'
#' @description Create a representation of a column break.
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
#' @title Page break for 'Word'
#' @description Object representing a line break
#' for a Word document. The result must be used
#' within a call to [fpar].
#' @inheritSection ftext usage
#' @examples
#' fp_t <- fp_text(font.size = 12, bold = TRUE)
#' an_fpar <- fpar("let's add a line break", run_linebreak(), ftext("and blah blah!", fp_t))
#'
#' x <- read_docx()
#' x <- body_add(x, an_fpar)
#' print(x, target = tempfile(fileext = ".docx"))
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

#' @export
#' @title Tab for 'Word'
#' @description Object representing a tab
#' in a Word document. The result must be used
#' within a call to [fpar]. It will only have
#' effects in Word output.
#'
#' Tabulation marks settings can be defined with [fp_tabs()] in
#' paragraph settings defined with [fp_par()].
#' @inheritSection ftext usage
#' @examples
#' z <- fp_tabs(
#'   fp_tab(pos = 0.5, style = "decimal"),
#'   fp_tab(pos = 1.5, style = "decimal")
#' )
#' par1 <- fpar(
#'   run_tab(), ftext("88."),
#'   run_tab(), ftext("987.45"),
#'   fp_p = fp_par(
#'     tabs = z
#'   )
#' )
#' par2 <- fpar(
#'   run_tab(), ftext("8."),
#'   run_tab(), ftext("670987.45"),
#'   fp_p = fp_par(
#'     tabs = z
#'   )
#' )
#' x <- read_docx()
#' x <- body_add(x, par1)
#' x <- body_add(x, par2)
#' print(x, target = tempfile(fileext = ".docx"))
#' @family run functions for reporting
run_tab <- function() {
  z <- list()
  class(z) <- c("run_tab", "run")

  z
}

#' @export
to_wml.run_tab <- function(x, add_ns = FALSE, ...) {
  out <- "<w:r><w:tab/></w:r>"
  out
}

# sections ----
inch_to_tweep <- function(x) round(x * 72 * 20, digits = 0)

#' @export
#' @title Page size object
#' @description The function creates a representation of the dimensions of a page. The
#' dimensions are defined by length, width and orientation. If the orientation is
#' in landscape mode then the length becomes the width and the width becomes the length.
#' @param width,height page width, page height (in inches).
#' @param orient page orientation, either 'landscape', either 'portrait'.
#' @param unit unit for width and height, one of "in", "cm", "mm".
#' @examples
#' page_size(orient = "landscape")
#' @family functions for section definition
page_size <- function(width = 21 / 2.54, height = 29.7 / 2.54, orient = "portrait", unit = "in") {

  width <- convin(unit = unit, x = width)
  height <- convin(unit = unit, x = height)

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
#' @title Section columns
#' @description The function creates a representation of the columns of a section.
#' @param widths columns widths in inches. If 3 values, 3 columns
#' will be produced.
#' @param space space in inches between columns.
#' @param sep if TRUE a line is separating columns.
#' @examples
#' section_columns()
#' @family functions for section definition
section_columns <- function(widths = c(2.5, 2.5), space = .25, sep = FALSE) {
  if (length(widths) < 2) {
    stop("length of widths should be at least 2")
  }

  z <- list(widths = widths, space = space, sep = sep)
  class(z) <- c("section_columns")
  z
}
#' @export
to_wml.section_columns <- function(x, add_ns = FALSE, ...) {
  widths <- x$widths * 20 * 72
  space <- x$space * 20 * 72

  columns_str_all_but_last <- sprintf(
    "<w:col w:w=\"%.0f\" w:space=\"%.0f\"/>",
    widths[-length(widths)], space
  )
  columns_str_last <- sprintf(
    "<w:col w:w=\"%.0f\"/>",
    widths[length(widths)]
  )
  columns_str <- c(columns_str_all_but_last, columns_str_last)

  sprintf(
    "<w:cols w:num=\"%.0f\" w:sep=\"%.0f\" w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
    length(widths), as.integer(x$sep),
    space, paste0(columns_str, collapse = "")
  )
}

#' @export
#' @title Page margins object
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
#' @title Section properties
#' @description A section is a grouping of blocks (ie. paragraphs and tables)
#' that have a set of properties that define pages on which the text will appear.
#'
#' A Section properties object stores information about page composition,
#' such as page size, page orientation, borders and margins.
#' @param page_size page dimensions, an object generated with function [page_size].
#' @param page_margins page margins, an object generated with function [page_mar].
#' @param type Section type. It defines how the contents of the section will be
#' placed relative to the previous section. Available types are "continuous"
#' (begins the section on the next paragraph), "evenPage" (begins on the next
#' even-numbered page), "nextColumn" (begins on the next column on the page),
#' "nextPage" (begins on the following page), "oddPage" (begins on the next
#' odd-numbered page).
#' @param section_columns section columns, an object generated with function [section_columns]. Use NULL (default value) for no content.
#' @param header_default content as a [block_list()] for the default page header. Use NULL (default value) for no content.
#' @param header_even content as a [block_list()] for the even page header. Use NULL (default value) for no content.
#' @param header_first content as a [block_list()] for the first page header. Use NULL (default value) for no content.
#' @param footer_default content as a [block_list()] for the default page footer. Use NULL (default value) for no content.
#' @param footer_even content as a [block_list()] for the even page footer. Use NULL (default value) for no content.
#' @param footer_first content as a [block_list()] for the default page footer. Use NULL (default value) for no content.
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
#' doc_1 <- body_add_table(doc_1, value = mtcars[1:10, ], style = "table_template")
#' doc_1 <- body_end_block_section(doc_1, value = landscape_one_column)
#' # there stops section with landscape_one_column
#'
#'
#' # there starts section with landscape_two_columns
#' doc_1 <- body_add_par(doc_1, value = paste(rep(letters, 50), collapse = " "))
#' doc_1 <- body_end_block_section(doc_1, value = landscape_two_columns)
#' # there stops section with landscape_two_columns
#'
#' doc_1 <- body_add_table(doc_1, value = mtcars[1:25, ], style = "table_template")
#'
#' print(doc_1, target = tempfile(fileext = ".docx"))
#'
#'
#' # an example with headers and footers -----
#' txt_lorem <- rep(
#'   "Purus lectus eros metus turpis mattis platea praesent sed. ",
#'   50
#' )
#' txt_lorem <- paste0(txt_lorem, collapse = "")
#'
#' header_first <- block_list(fpar(ftext("text for first page header")))
#' header_even <- block_list(fpar(ftext("text for even page header")))
#' header_default <- block_list(fpar(ftext("text for default page header")))
#' footer_first <- block_list(fpar(ftext("text for first page footer")))
#' footer_even <- block_list(fpar(ftext("text for even page footer")))
#' footer_default <- block_list(fpar(ftext("text for default page footer")))
#'
#' ps <- prop_section(
#'   header_default = header_default, footer_default = footer_default,
#'   header_first = header_first, footer_first = footer_first,
#'   header_even = header_even, footer_even = footer_even
#' )
#' x <- read_docx()
#' for (i in 1:20) {
#'   x <- body_add_par(x, value = txt_lorem)
#' }
#' x <- body_set_default_section(
#'   x,
#'   value = ps
#' )
#' print(x, target = tempfile(fileext = ".docx"))
#' @seealso [block_section]
#' @family functions for section definition
#' @section Illustrations:
#'
#' \if{html}{\figure{prop_section_doc_1.png}{options: width=80\%}}
prop_section <- function(page_size = NULL, page_margins = NULL,
                         type = NULL, section_columns = NULL,
                         header_default = NULL, header_even = NULL, header_first = NULL,
                         footer_default = NULL, footer_even = NULL, footer_first = NULL) {
  z <- list()

  if (!is.null(header_default) && !inherits(header_default, "block_list")) {
    stop("header_default must be NULL or an object generated by `block_list()`.")
  }
  z$header_default <- header_default

  if (!is.null(header_even) && !inherits(header_even, "block_list")) {
    stop("header_even must be NULL or an object generated by `block_list()`.")
  }
  z$header_even <- header_even

  if (!is.null(header_first) && !inherits(header_first, "block_list")) {
    stop("header_first must be NULL or an object generated by `block_list()`.")
  }
  z$header_first <- header_first

  if (!is.null(footer_default) && !inherits(footer_default, "block_list")) {
    stop("footer_default must be NULL or an object generated by `block_list()`.")
  }
  z$footer_default <- footer_default

  if (!is.null(footer_even) && !inherits(footer_even, "block_list")) {
    stop("footer_even must be NULL or an object generated by `block_list()`.")
  }
  z$footer_even <- footer_even

  if (!is.null(footer_first) && !inherits(footer_first, "block_list")) {
    stop("footer_first must be NULL or an object generated by `block_list()`.")
  }
  z$footer_first <- footer_first

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
    "<w:sectPr w:officer=\"true\"",
    if (add_ns) " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
    ">",
    if (!is.null(x$header_default)) {
      paste0(
        "<w:headerReference w:type=\"default\">",
        "<w:hdr xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\">",
        to_wml(x$header_default),
        "</w:hdr>",
        "</w:headerReference>"
      )
    },
    if (!is.null(x$header_even)) {
      paste0(
        "<w:headerReference w:type=\"even\">",
        "<w:hdr xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\">",
        to_wml(x$header_even),
        "</w:hdr>",
        "</w:headerReference>"
      )
    },
    if (!is.null(x$header_first)) {
      paste0(
        "<w:headerReference w:type=\"first\">",
        "<w:hdr xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\">",
        to_wml(x$header_first),
        "</w:hdr>",
        "</w:headerReference>"
      )
    },
    if (!is.null(x$footer_default)) {
      paste0(
        "<w:footerReference w:type=\"default\">",
        "<w:ftr xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\">",
        to_wml(x$footer_default),
        "</w:ftr>",
        "</w:footerReference>"
      )
    },
    if (!is.null(x$footer_even)) {
      paste0(
        "<w:footerReference w:type=\"even\">",
        "<w:ftr xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\">",
        to_wml(x$footer_even),
        "</w:ftr>",
        "</w:footerReference>"
      )
    },
    if (!is.null(x$footer_first)) {
      paste0(
        "<w:footerReference w:type=\"first\">",
        "<w:ftr xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\">",
        to_wml(x$footer_first),
        "</w:ftr>",
        "</w:footerReference>"
      )
    },
    if (!is.null(x$page_margins)) to_wml(x$page_margins),
    if (!is.null(x$page_size)) to_wml(x$page_size),
    if (!is.null(x$type)) paste0("<w:type w:val=\"", x$type, "\"/>"),
    if (!is.null(x$section_columns)) to_wml(x$section_columns) else "<w:cols/>",
    "</w:sectPr>"
  )
}


# external_img ----
#' @export
#' @title External image
#' @description Wraps an image in an object that can then be embedded
#' in a PowerPoint slide or within a Word paragraph.
#'
#' The image is added as a shape in PowerPoint (it is not possible to mix text and
#' images in a PowerPoint form). With a Word document, the image will be
#' added inside a paragraph.
#' @param src image file path
#' @param width,height size of the image file. It can be ignored
#' if parameter `guess_size=TRUE`, see parameter `guess_size`.
#' @param guess_size If package 'magick' is installed, this option
#' can be used (set it to `TRUE`). The images will be read and
#' width and height will be guessed.
#' @param unit unit for width and height, one of "in", "cm", "mm".
#' @param alt alternative text for images
#' @inheritSection ftext usage
#' @examples
#' # wrap r logo with external_img ----
#' srcfile <- file.path(R.home("doc"), "html", "logo.jpg")
#' extimg <- external_img(
#'   src = srcfile, height = 1.06 / 2,
#'   width = 1.39 / 2
#' )
#'
#' # pptx example ----
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(
#'   x = doc, value = extimg,
#'   location = ph_location_type(type = "body"),
#'   use_loc_size = FALSE
#' )
#' print(doc, target = tempfile(fileext = ".pptx"))
#'
#' fp_t <- fp_text(font.size = 20, color = "red")
#' an_fpar <- fpar(extimg, ftext(" is cool!", fp_t))
#'
#' # docx example ----
#' x <- read_docx()
#' x <- body_add(x, an_fpar)
#' print(x, target = tempfile(fileext = ".docx"))
#' @seealso [ph_with], [body_add], [fpar]
#' @family run functions for reporting
external_img <- function(src, width = .5, height = .2, unit = "in", guess_size = FALSE, alt = "") {
  # note: should it be vectorized
  check_src <- all(grepl("^rId", src)) || all(file.exists(src))
  if (!check_src) {
    stop("src must be a string starting with 'rId' or an existing image filename")
  }

  width <- convin(unit = unit, x = width)
  height <- convin(unit = unit, x = height)

  if (length(src) > 1) {
    if (length(width) == 1) width <- rep(width, length(src))
    if (length(height) == 1) height <- rep(height, length(src))
  }

  if (guess_size) {
    if (!requireNamespace("magick", quietly = TRUE)) {
      stop("package magick is required when using `guess_size` option.")
    }
    sizes <- lapply(src, function(x) {
      if (grepl("\\.svg$", x)) {
        z <- magick::image_read_svg(x)
      } else {
        z <- magick::image_read(x)
      }
      z <- magick::image_data(z)
      attr(z, "dim")[-1]
    })
    sizes <- do.call(rbind, sizes)
    width <- sizes[, 1] / 72
    height <- sizes[, 2] / 72
  }

  class(src) <- c("external_img", "cot", "run")
  attr(src, "dims") <- list(width = width, height = height)
  attr(src, "alt") <- alt
  src
}


#' @export
dim.external_img <- function(x) {
  x <- attr(x, "dims")
  data.frame(width = x$width, height = x$height)
}


#' @export
as.data.frame.external_img <- function(x, ...) {
  dimx <- attr(x, "dims")
  data.frame(path = as.character(x), width = dimx$width, height = dimx$height, alt = attr(x, "alt"), stringsAsFactors = FALSE)
}

temp_blipfill <- function(value, ns = "p") {
  svg_src <- NULL
  if (grepl("\\.svg", value)) {
    if (!requireNamespace("rsvg")) {
      stop("package 'rsvg' is required to convert svg file to rasters.")
    }

    svg_src <- tempfile(fileext = ".svg")
    file.copy(as.character(value), to = svg_src)

    img_src <- tempfile(fileext = ".png")
    rsvg::rsvg_png(svg_src, file = img_src)
  } else {
    img_src <- tempfile(fileext = gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", as.character(value)))
    file.copy(as.character(value), to = img_src)
  }

  if (!is.null(svg_src)) {
    blipfill <- paste0(
      "<", ns, ":blipFill>",
      sprintf("<a:blip cstate=\"print\" r:embed=\"%s\">", img_src),
      "<a:extLst>",
      "<a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">",
      sprintf("<asvg:svgBlip r:embed=\"%s\" xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\"/>", svg_src),
      "</a:ext>",
      "</a:extLst>",
      "</a:blip>",
      "<a:stretch><a:fillRect/></a:stretch>",
      "</", ns, ":blipFill>"
    )
  } else {
    blipfill <- paste0(
      "<", ns, ":blipFill>",
      sprintf("<a:blip cstate=\"print\" r:embed=\"%s\"/>", img_src),
      "<a:stretch><a:fillRect/></a:stretch>",
      "</", ns, ":blipFill>"
    )
  }
  blipfill
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

  blipfill <- temp_blipfill(x, ns = "pic")

  paste0(
    open_tag,
    "<w:rPr/><w:drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">",
    sprintf("<wp:extent cx=\"%.0f\" cy=\"%.0f\"/>", width * 12700 * 72, height * 12700 * 72),
    "<wp:docPr id=\"\" name=\"\" descr=\"", attr(x, "alt"), "\"/>",
    "<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>",
    "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
    "<pic:nvPicPr>",
    "<pic:cNvPr id=\"\" name=\"\"/>",
    "<pic:cNvPicPr><a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>",
    "</pic:cNvPicPr></pic:nvPicPr>",
    blipfill,
    "<pic:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/>",
    sprintf("<a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></pic:spPr>", width * 12700, height * 12700),
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
  )
}


#' @export
to_pml.external_img <- function(x, add_ns = FALSE,
                                left = 0, top = 0,
                                width = 3, height = 3,
                                bg = "transparent",
                                rot = 0,
                                ph = "<p:ph/>",
                                label = "",
                                ln = sp_line(lwd = 0, linecmpd = "solid", lineend = "rnd"),
                                ...) {
  if (!is.null(bg) && !is.color(bg)) {
    stop("bg must be a valid color.", call. = FALSE)
  }

  bg_str <- solid_fill_pml(bg)
  ln_str <- ln_pml(ln)

  xfrm_str <- a_xfrm_str(left = left, top = top, width = width, height = height, rot = rot)
  if (is.null(ph) || is.na(ph)) {
    ph <- "<p:ph/>"
  }
  blipfill <- temp_blipfill(x, ns = "p")
  id <- uuid_generate()
  str <- "
<p:pic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">
  <p:nvPicPr>
    <p:cNvPr id=\"%s\" name=\"%s\" descr=\"%s\"/>
    <p:cNvPicPr><a:picLocks noGrp=\"1\"/></p:cNvPicPr>
    <p:nvPr>%s</p:nvPr>
  </p:nvPicPr>
  %s
  <p:spPr>%s<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>%s%s</p:spPr>
</p:pic>
"
  sprintf(str, id, label, attr(x, "alt"), ph, blipfill, xfrm_str, bg_str, ln_str)
}



#' @export
to_html.external_img <- function(x, ...) {
  dims <- attr(x, "dims")
  width <- dims$width
  height <- dims$height
  b64_str <- image_to_base64(as.character(x))
  sprintf("<img style=\"vertical-align:middle;width:%.0fpx;height:%.0fpx;\" src=\"%s\" alt=\"%s\"/>", width * 72, height * 72, b64_str, attr(x, "alt"))
}

# run_footnoteref ----
#' @export
#' @title Word footnote reference
#' @description Wraps a footnote reference in an object that can then be inserted
#' as a run/chunk with [fpar()] or within an R Markdown document.
#' @param prop formatting text properties returned by
#' [fp_text_lite()] or [fp_text()]. It also can be NULL in
#' which case, no formatting is defined (the default is applied).
#' @family run functions for reporting
#' @examples
#' run_footnoteref()
#' to_wml(run_footnoteref())
run_footnoteref <- function(prop = NULL) {
  z <- list(pr = prop)
  class(z) <- c("run_footnoteref", "run")
  z
}
#' @export
to_wml.run_footnoteref <- function(x, add_ns = FALSE, ...) {
  open_tag <- wr_ns_no
  if (add_ns) {
    open_tag <- wr_ns_yes
  }

  str <- paste0(
    open_tag,
    if (!is.null(x$pr)) rpr_wml(x$pr),
    "<w:footnoteRef/></w:r>"
  )

  str
}

# run_footnote ----
#' @export
#' @title Footnote for 'Word'
#' @description Wraps a footnote in an object that can then be inserted
#' as a run/chunk with [fpar()] or within an R Markdown document.
#' @param x a set of blocks to be used as footnote content returned by
#'   function [block_list()].
#' @param prop formatting text properties returned by
#' [fp_text_lite()] or [fp_text()]. It also can be NULL in
#' which case, no formatting is defined (the default is applied).
#' @family run functions for reporting
#' @examples
#' library(officer)
#'
#' fp_bold <- fp_text_lite(bold = TRUE)
#' fp_refnote <- fp_text_lite(vertical.align = "superscript")
#'
#' img.file <- file.path(R.home("doc"), "html", "logo.jpg")
#' bl <- block_list(
#'   fpar(ftext("hello", fp_bold)),
#'   fpar(
#'     ftext("hello world", fp_bold),
#'     external_img(src = img.file, height = 1.06, width = 1.39)
#'   )
#' )
#'
#' a_par <- fpar(
#'   "this paragraph contains a note ",
#'   run_footnote(x = bl, prop = fp_refnote),
#'   "."
#' )
#'
#' doc <- read_docx()
#' doc <- body_add_fpar(doc, value = a_par, style = "Normal")
#'
#' print(doc, target = tempfile(fileext = ".docx"))
run_footnote <- function(x, prop = NULL) {
  z <- list(footnote = x, pr = prop)
  class(z) <- c("run_footnote", "run")
  z
}
#' @export
to_wml.run_footnote <- function(x, add_ns = FALSE, ...) {
  open_tag <- wr_ns_no
  if (add_ns) {
    open_tag <- wr_ns_yes
  }

  x$footnote[[1]]$chunks <- append(list(run_footnoteref(x$pr)), x$footnote[[1]]$chunks)
  blocks <- sapply(x$footnote, to_wml)
  blocks <- paste(blocks, collapse = "")

  id <- basename(tempfile(pattern = "footnote"))
  base_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""
  footnote_xml <- paste0(
    "<w:footnote ", base_ns, " w:id=\"",
    id,
    "\">",
    blocks, "</w:footnote>"
  )

  footnote_ref_xml <- paste0(
    open_tag,
    if (!is.null(x$pr)) rpr_wml(x$pr),
    "<w:footnoteReference w:id=\"",
    id,
    "\">",
    footnote_xml,
    "</w:footnoteReference>",
    "</w:r>"
  )

  footnote_ref_xml
}

# run_comment ----
#' @export
#' @title Comment for 'Word'
#' @description Add a comment on a run object.
#' @param cmt a set of blocks to be used as comment content returned by
#'   function [block_list()].
#' the "run functions for reporting".
#' @param author comment author.
#' @param date comment date
#' @param initials comment initials
#' @param prop formatting text properties returned by
#' [fp_text_lite()] or [fp_text()]. It also can be NULL in
#' which case, no formatting is defined (the default is applied).
#' @param run a run object, made with a call to one of
#' @examples
#' fp_bold <- fp_text_lite(bold = TRUE)
#' fp_red <- fp_text_lite(color = "red")
#'
#' bl <- block_list(
#'   fpar(ftext("Comment multiple words.", fp_bold)),
#'   fpar(
#'     ftext("Second line.", fp_red)
#'   )
#' )
#'
#' comment1 <- run_comment(
#'   cmt = bl,
#'   run = ftext("with a comment"),
#'   author = "Author Me",
#'   date = Sys.Date(),
#'   initials = "AM"
#' )
#' par1 <- fpar("A paragraph ", comment1)
#'
#' bl <- block_list(
#'   fpar(ftext("Comment a paragraph."))
#' )
#'
#' comment2 <- run_comment(
#'   cmt = bl, run = ftext("A commented paragraph"),
#'   author = "Author You",
#'   date = Sys.Date(),
#'   initials = "AY"
#' )
#' par2 <- fpar(comment2)
#'
#' doc <- read_docx()
#' doc <- body_add_fpar(doc, value = par1, style = "Normal")
#' doc <- body_add_fpar(doc, value = par2, style = "Normal")
#'
#' print(doc, target = tempfile(fileext = ".docx"))
#' @family run functions for reporting
run_comment <- function(cmt, run = ftext(""), author = "", date = "", initials = "", prop = NULL) {
  all_run <- FALSE

  if (inherits(run, "run")) {
    run <- list(run)
  }

  if (is.list(run) && !inherits(run, "run")) {
    all_run <- all(vapply(run, inherits, FUN.VALUE = FALSE, what = "run"))
  }

  if (!all_run) {
    stop("`run` must be a run object (ftext for example) or a list of run objects.")
  }

  z <- list(comment = cmt, run = run, author = author, date = date, initials = initials, pr = prop)
  class(z) <- c("run_comment", "run")
  z
}

#' @export
to_wml.run_comment <- function(x, add_ns = FALSE, ...) {
  runs <- lapply(x$run, to_wml, add_ns = add_ns, ...)
  runs <- do.call(paste0, runs)

  open_tag <- wr_ns_no
  if (add_ns) {
    open_tag <- wr_ns_yes
  }

  blocks <- sapply(x$comment, to_wml)
  blocks <- paste(blocks, collapse = "")

  id <- basename(tempfile(pattern = "comment"))
  base_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""
  comment_xml <- paste0(
    sprintf(
      "<w:comment %s  w:id=\"%s\" w:author=\"%s\" w:date=\"%s\" w:initials=\"%s\">",
      base_ns, id, x$author, x$date, x$initials
    ),
    blocks,
    "</w:comment>"
  )

  comment_ref_xml <- paste0(
    "<w:commentRangeStart w:id=\"", id, "\"/>",
    runs,
    "<w:commentRangeEnd w:id=\"", id, "\"/>",
    open_tag,
    if (!is.null(x$pr)) rpr_wml(x$pr),
    "<w:commentReference w:id=\"",
    id,
    "\">",
    comment_xml,
    "</w:commentReference>",
    "</w:r>"
  )

  comment_ref_xml
}
