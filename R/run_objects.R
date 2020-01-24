# generics -----

#' @export
#' @title Convert officer objects to WordprocessingML
#' @description Convert an object made with package officer
#' to WordprocessingML. The result is a string.
#' @param x a location for a placeholder.
#' @param ... Arguments to be passed to methods
to_wml <- function(x, ...) {
  UseMethod("to_wml")
}

bookmark <- function(id, str) {
  new_id <- UUIDgenerate()
  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\"/>", new_id, id)
  bm_start_end <- sprintf("<w:bookmarkEnd w:id=\"%s\"/>", new_id)
  paste0(bm_start_str, str, bm_start_end)
}

# seqfield ----
#' @export
#' @title seqfield
#' @description Create a seqfield
#' @param seqfield seqfield string
run_seqfield <- function(seqfield) {
  z <- list(
    seqfield = seqfield
  )
  class(z) <- c("run_seqfield", "run")
  z
}


#' @export
to_wml.run_seqfield <- function(x, ...) {
  xml_elt_1 <- paste0(
    wml_with_ns("w:r"),
    "<w:rPr/>",
    "<w:fldChar w:fldCharType=\"begin\"/>",
    "</w:r>"
  )
  xml_elt_2 <- paste0(
    wml_with_ns("w:r"),
    "<w:rPr/>",
    sprintf("<w:instrText xml:space=\"preserve\">%s</w:instrText>", x$seqfield),
    "</w:r>"
  )
  xml_elt_3 <- paste0(
    wml_with_ns("w:r"),
    "<w:rPr/>",
    "<w:fldChar w:fldCharType=\"end\"/>",
    "</w:r>"
  )
  out <- paste0(xml_elt_1, xml_elt_2, xml_elt_3)

  out
}

# autonum ----

#' @export
#' @title auto number
#' @description Create a string representation of a number
#' @param seq_id sequence identifier
#' @param pre_label,post_label text to add before and after number
#' @examples
#' run_autonum()
#' run_autonum(seq_id = "fig", pre_label = "fig. ")
run_autonum <- function(seq_id = "table", pre_label = "TABLE ", post_label = ": ") {
  z <- list(
    seq_id = seq_id,
    pre_label = pre_label,
    post_label = post_label
  )
  class(z) <- c("run_autonum", "run")

  z
}

#' @export
to_wml.run_autonum <- function(x, ...) {
  run_str_pre <- sprintf("<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", x$pre_label)
  run_str_post <- sprintf("<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", x$post_label)
  sqf <- run_seqfield(seqfield = paste0("SEQ ", x$seq_id, " \u005C* Arabic \u005Cs 1 \u005C* MERGEFORMAT"))
  sf_str <- to_wml(sqf)
  out <- paste0(run_str_pre, sf_str, run_str_post)

  out
}

# reference ----

#' @export
#' @title reference
#' @description Create a representation of a reference
#' @param id reference id, a string
#' @examples
#' run_reference('a_ref')
run_reference <- function(id) {
  z <- paste0(" REF ", id, " \\h ")

  z <- list(
    id = id
  )
  class(z) <- c("run_reference", "run")

  z
}

#' @export
to_wml.run_reference <- function(x, ...) {
  out <- to_wml(run_seqfield(seqfield = x$id))
  out
}

# breaks ----

#' @export
#' @export
#' @title page break
#' @description Create a representation of a page break
#' @examples
#' run_pagebreak()
run_pagebreak <- function() {
  z <- list()
  class(z) <- c("run_pagebreak", "run")

  z
}

#' @export
to_wml.run_pagebreak <- function(x, ...) {
  out <- "<w:r><w:br w:type=\"page\"/></w:r>"
  out
}

#' @export
#' @title column break
#' @description Create a representation of a column break
#' @examples
#' run_columnbreak()
run_columnbreak <- function() {
  z <- list()
  class(z) <- c("run_columnbreak", "run")

  z
}

#' @export
to_wml.run_columnbreak <- function(x, ...) {
  out <- "<w:r><w:br w:type=\"column\"/></w:r>"
  out
}

#' @export
#' @title line break
#' @description Create a representation of a line break
#' @examples
#' run_linebreak()
run_linebreak <- function() {
  z <- list()
  class(z) <- c("run_linebreak", "run")

  z
}

#' @export
to_wml.run_linebreak <- function(x, ...) {
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
page_size <- function(width = 21 / 2.54, height = 29.7 / 2.54, orient = "portrait") {
  if (orient %in% "portrait") {
    h <- height
    w <- width
  } else {
    w <- height
    h <- width
  }

  z <- list(width = w, height = h, orient = orient)
  class(z) <- c("page_size")
  z
}

#' @export
to_wml.page_size <- function(x, ...) {
  out <- sprintf(
    "<w:pgSz w:h=\"%.0f\" w:w=\"%.0f\" w:orient=\"%s\"/>",
    inch_to_tweep(x$height), inch_to_tweep(x$width), x$orient
  )
  out
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
page_mar <- function(bottom = 1, top = 1, right = 1, left = 1,
                     header = 0.5, footer = 0.5, gutter = 0.5) {
  z <- list(header = header, bottom = bottom, top = top, right = right, left = left, footer = footer, gutter = gutter)
  class(z) <- c("page_mar")
  z
  # <w:pgMar w:header="720" w:bottom="1440" w:top="1440" w:right="2880" w:left="2880" w:footer="720" w:gutter="720"/>
}
#' @export
to_wml.page_mar <- function(x, ...) {
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
#' @description A section is a
#' grouping of blocks (ie. paragraphs and tables) that have a set of properties
#' that define pages on which the text will appear.
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
#' @examples
#' prop_section(
#'   page_size = page_size(orient = "landscape"),
#'   page_margins = page_mar(top = 2),
#'   type = "continuous")
#' @seealso [block_section]
prop_section <- function(page_size, page_margins, type) {
  if (!type %in% docx_section_type) {
    stop("type must be one of ", paste(shQuote(docx_section_type), collapse = ", "))
  }
  if (!inherits(page_size, "page_size")) {
    stop("page_size must be a page_size object")
  }
  if (!inherits(page_margins, "page_mar")) {
    stop("page_margins must be a page_mar object")
  }

  z <- list(page_size = page_size, page_margins = page_margins, type = type)
  class(z) <- c("prop_section", "prop")

  z
}

#' @export
to_wml.prop_section <- function(x, ...) {
  out <- sprintf(
    "<w:sectPr>%s%s<w:type w:val=\"%s\"/></w:sectPr>",
    to_wml(x$page_margins),
    to_wml(x$page_size),
    x$type
  )
  out
}
