# generics -----

#' @export
#' @title Convert officer objects to WordprocessingML
#' @description Convert an object made with package officer
#' to WordprocessingML. The result is a string.
#' @param x object
#' @param add_ns should namespace be added to the top tag
#' @param ... Arguments to be passed to methods
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
to_pml <- function(x, add_ns = FALSE, ...) {
  UseMethod("to_pml")
}

#' @export
#' @title Convert officer objects to HTML
#' @description Convert an object made with package officer
#' to HTML. The result is a string.
#' @param x object
#' @param ... Arguments to be passed to methods
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
to_wml.run_seqfield <- function(x, add_ns = FALSE, ...) {
  xml_elt_1 <- paste0(
    wr_ns_yes,
    "<w:rPr/>",
    "<w:fldChar w:fldCharType=\"begin\"/>",
    "</w:r>"
  )
  xml_elt_2 <- paste0(
    wr_ns_yes,
    "<w:rPr/>",
    sprintf("<w:instrText xml:space=\"preserve\">%s</w:instrText>", x$seqfield),
    "</w:r>"
  )
  xml_elt_3 <- paste0(
    wr_ns_yes,
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
to_wml.run_autonum <- function(x, add_ns = FALSE, ...) {
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
to_wml.run_reference <- function(x, add_ns = FALSE, ...) {
  out <- to_wml(run_seqfield(seqfield = x$id))
  out
}

# breaks ----

#' @export
#' @title page break for Word
#' @description Object representing a page break for a Word document. The
#' result must be used within a call to [fpar].
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
#' @examples
#' run_columnbreak()
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
to_wml.page_size <- function(x, add_ns = FALSE, ...) {
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
to_wml.prop_section <- function(x, add_ns = FALSE, ...) {
  out <- sprintf(
    "<w:sectPr>%s%s<w:type w:val=\"%s\"/></w:sectPr>",
    to_wml(x$page_margins),
    to_wml(x$page_size),
    x$type
  )
  out
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
#' @examples
#'
#' @example examples/external_img.R
#' @seealso [ph_with], [body_add], [fpar]
#' @family run functions for reporting
external_img <- function(src, width = .5, height = .2) {
  # note: should it be vectorized
  check_src <- all(grepl("^rId", src)) || all(file.exists(src))
  if( !check_src ){
    stop("src must be a string starting with 'rId' or an image filename")
  }

  class(src) <- c("external_img", "cot")
  attr(src, "dims") <- list(width = width, height = height)
  src
}

dim.external_img <- function( x ){
  x <- attr(x, "dims")
  data.frame(width = x$width, height = x$height)
}

as.data.frame.external_img <- function( x, ... ){
  dimx <- attr(x, "dims")
  data.frame(path = as.character(x), width = dimx$width, height = dimx$height)
}


pic_pml <- function( left = 0, top = 0, width = 3, height = 3,
                     bg = "transparent", rot = 0, label = "", ph = "<p:ph/>", src){

  if( !is.null(bg) && !is.color( bg ) )
    stop("bg must be a valid color.", call. = FALSE )

  bg_str <- solid_fill_pml(bg)

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
    <p:cNvPr id=\"0\" name=\"%s\"/>
    <p:cNvPicPr><a:picLocks noGrp=\"1\"/></p:cNvPicPr>
    <p:nvPr>%s</p:nvPr>
  </p:nvPicPr>
  %s
  <p:spPr>%s%s</p:spPr>
</p:pic>
"
  sprintf(str, label, ph, blipfill, xfrm_str, bg_str )

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
         "<wp:docPr id=\"\" name=\"\"/>",
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


# ftext -----
#' @export
#' @title formatted chunk of text
#' @description Format a chunk of text with text formatting properties (bold, color, ...).
#'
#' The function allows you to create pieces of text formatted in a certain way.
#' You should use this function in conjunction with [fpar] to create paragraphs
#' consisting of differently formatted text parts.
#' @param text text value, a string.
#' @param prop formatting text properties returned by [fp_text].
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
ftext <- function(text, prop) {
  out <- list( value = formatC(text), pr = prop )
  class(out) <- c("ftext", "cot")
  out
}

#' @export
print.ftext <- function (x, ...) {
  cat("text: ", x$value, "\nformat:\n", sep = "")
  print(x$pr)
}

#' @export
to_wml.ftext <- function(x, add_ns = FALSE, ...) {
  out <- paste0("<w:r>", rpr_wml(x$pr),
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
  paste0(open_tag, rpr_pml(x$pr),
         "<a:t>", htmlEscapeCopy(x$value), "</a:t></a:r>")
}

#' @export
to_html.ftext <- function(x, ...) {
  sprintf("<span style=\"%s\">%s</span>", rpr_css(x$pr), htmlEscapeCopy(x$value))
}


