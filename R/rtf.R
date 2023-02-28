# to_rtf ----
#' @export
#' @title Convert officer objects to RTF
#' @description Convert an object made with package officer
#' to RTF.
#' @param x object to convert to RTF. Supported
#' objects are:
#' - [ftext()]
#' - [external_img()]
#' - [run_word_field()]
#' - [run_pagebreak()]
#' - [run_columnbreak()]
#' - [run_linebreak()]
#' - [hyperlink_ftext()]
#' - [block_list()]
#' - [fpar()]
#' @param ... Arguments to be passed to methods
#' @family functions for officer extensions
#' @return a string containing the RTF code
#' @keywords internal
to_rtf <- function(x, ...) {
  UseMethod("to_rtf")
}

## to_rtf blocks ----
#' @export
to_rtf.fpar <- function(x, ...) {
  chks <- fortify_fpar(x)
  rtf_chunks <- lapply(chks, to_rtf)
  rtf_chunks$collapse <- ""
  rtf_chunks <- do.call(paste0, rtf_chunks)
  paste0("{\\pard", ppr_rtf(x$fp_p), rtf_chunks, "\\par}")
}

#' @export
to_rtf.default <- function(x, ...) {
  if (is.null(x)) {
    return("")
  }
  as.character(x)
}

#' @export
to_rtf.block_list <- function(x, ...) {
  rtf_blocks <- lapply(x, to_rtf)
  rtf_blocks$collapse <- ""
  do.call(paste0, rtf_blocks)
}



## to_rtf chunks ----

#' @export
to_rtf.ftext <- function(x, ...) {
  rtf_fp_strings <- ""
  value <- str_encode_to_rtf(x$value)
  if (!is.null(x$pr)) {
    rtf_fp_strings <- rpr_rtf(x$pr)
    if (!is.na(x$pr$shading.color)) {
      value <- paste0(
        "%ftshading:",
        x$pr$shading.color,
        "%",
        value,
        "\\highlight0"
      )
    }
  }

  paste0(rtf_fp_strings, value)
}

#' @export
to_rtf.external_img <- function(x, ...) {
  dims <- attr(x, "dims")
  width <- dims$width
  height <- dims$height

  if (!grepl("\\.png$", x)) {
    imgpath <- tempfile(fileext = ".png")
    img <- magick::image_read(x)
    magick::image_write(img, path = imgpath)
  } else {
    imgpath <- x
  }

  # https://stat.ethz.ch/pipermail/r-help/2017-September/449253.html
  dat <- readBin(imgpath, what = "raw", size = 1, endian = "little", n = 1e+8)
  dat <- paste(dat, collapse = "")

  rtf_str <- sprintf(
    "{\\pict\\pngblip\\picwgoal%.0f\\pichgoal%.0f ",
    width * 1440, height * 1440
  )
  paste(rtf_str, dat, "\n}", sep = "")
}

#' @export
to_rtf.run_word_field <- function(x, ...) {
  rtf_fp <- ""
  if (!is.null(x$pr)) {
    rtf_fp <- rpr_rtf(x$pr)
  }
  sprintf("%s{\\field{\\*\\fldinst %s}}", rtf_fp, x$field)
}

#' @export
to_rtf.run_pagebreak <- function(x, ...) {
  "\\page"
}
#' @export
to_rtf.run_columnbreak <- function(x, ...) {
  "\\column"
}
#' @export
to_rtf.run_linebreak <- function(x, ...) {
  "\\line"
}

#' @export
to_rtf.hyperlink_ftext <- function(x, ...) {
  rtf_fp <- ""
  if (!is.null(x$pr)) {
    rtf_fp <- rpr_rtf(x$pr)
  }
  sprintf(
    "%s{\\field{\\*\\fldinst HYPERLINK \"%s\"}{\\fldrslt %s}}",
    rtf_fp,
    officer_url_encode(x$href),
    str_encode_to_rtf(x$value)
  )
}
rtf_bookmark <- function(id, str) {
  bm_start_str <- sprintf("{\\*\\bkmkstart %s}", id)
  bm_start_end <- sprintf("{\\*\\bkmkend %s}", id)
  paste0(bm_start_str, str, bm_start_end)
}

#' @export
to_rtf.run_autonum <- function(x, ...) {
  if(is.null(x$seq_id)) return("")

  run_str_pre <- to_rtf(ftext(x$pre_label, prop = x$pr))
  run_str_post <- to_rtf(ftext(x$post_label, prop = x$pr))

  seq_str <- paste0("SEQ ", x$seq_id, " \\\\* Arabic")
  if(!is.null(x$start_at) && is.numeric(x$start_at)){
    seq_str <- paste(seq_str, "\\r", as.integer(x$start_at))
  }

  sqf <- run_word_field(field = seq_str, prop = x$pr)
  sf_str <- to_rtf(sqf)

  if(x$tnd > 0){
    z <- paste0(
      to_rtf(run_word_field(field = paste0("STYLEREF ", x$tnd, " \\r"), prop = x$pr)),
      to_rtf(ftext(x$tns, prop = x$pr))
    )
    sf_str <- paste0(z, sf_str)
  }

  if(!is.null(x$bookmark)){
    if(x$bookmark_all){
      out <- paste0(
        rtf_bookmark(id = x$bookmark, str = paste0(run_str_pre, sf_str)),
        run_str_post)
    } else {
      sf_str <- rtf_bookmark(id = x$bookmark, sf_str)
      out <- paste0(run_str_pre, sf_str, run_str_post)
    }
  } else {
    out <- paste0(run_str_pre, sf_str, run_str_post)
  }

  out
}

#' @export
to_rtf.run_reference <- function(x, ...) {
  out <- to_rtf(run_word_field(field = paste0(" REF ", x$id, " \\\\h "), prop = x$pr))
  out
}

## to_rtf sections ----

#' @export
to_rtf.block_section <- function(x, ...) {
  paste0("\\sect", to_rtf(x$property))
}

#' @export
to_rtf.prop_section <- function(x, ...) {
  rtf_type <- if (is.null(x$type)) {
    ""
  } else if (x$type %in% "nextPage") {
    "\\sbkpage"
  } else if (x$type %in% "oddPage") {
    "\\sbkodd"
  } else if (x$type %in% "evenPage") {
    "\\sbkeven"
  } else {
    "\\sbknone"
  }

  hfr_type <- guess_hfr_type(x)
  extra_rtf <- ""
  if ("none" %in% hfr_type) {
    hfr_str <- ""
  } else if ("default" %in% hfr_type) {
    hfr_str <- paste0(
      if (!is.null(x$header_default)) "{\\header ", to_rtf(x$header_default), "}",
      if (!is.null(x$footer_default)) "{\\footer ", to_rtf(x$footer_default), "}",
      ""
    )
  } else if ("evenodd" %in% hfr_type) {
    extra_rtf <- "\\facingp"
    hfr_str <- paste0(
      if (!is.null(x$header_default)) "{\\headerl ", to_rtf(x$header_default), "}",
      if (!is.null(x$header_even)) "{\\headerr ", to_rtf(x$header_even), "}",
      if (!is.null(x$footer_default)) "{\\footerl ", to_rtf(x$footer_default), "}",
      if (!is.null(x$footer_even)) "{\\footerr ", to_rtf(x$footer_even), "}",
    )
  } else if ("evenoddfirst" %in% hfr_type) {
    extra_rtf <- "\\facingp\\titlepg"
    hfr_str <- paste0(
      if (!is.null(x$header_default)) "{\\headerl ", to_rtf(x$header_default), "}",
      if (!is.null(x$header_even)) "{\\headerr ", to_rtf(x$header_even), "}",
      if (!is.null(x$header_first)) "{\\headerf ", to_rtf(x$header_first), "}",
      if (!is.null(x$footer_default)) "{\\footerl ", to_rtf(x$footer_default), "}",
      if (!is.null(x$footer_even)) "{\\footerr ", to_rtf(x$footer_even), "}",
      if (!is.null(x$footer_first)) "{\\footerf ", to_rtf(x$footer_first), "}",
      ""
    )
  } else {
    extra_rtf <- "\\titlepg"
    hfr_str <- paste0(
      if (!is.null(x$header_default)) "{\\header ", to_rtf(x$header_default), "}",
      if (!is.null(x$header_first)) "{\\headerf ", to_rtf(x$header_first), "}",
      if (!is.null(x$footer_default)) "{\\footer ", to_rtf(x$footer_default), "}",
      if (!is.null(x$footer_first)) "{\\footerf ", to_rtf(x$footer_first), "}"
    )
  }

  rtf_str <- paste0(
    if (!is.null(x$page_margins)) to_rtf(x$page_margins),
    if (!is.null(x$page_size)) to_rtf(x$page_size),
    rtf_type,
    if (!is.null(x$section_columns)) to_rtf(x$section_columns),
    extra_rtf
  )

  # sections def and hdr/ftr content are separated in rtf
  attr(rtf_str, "hfr") <- hfr_str

  rtf_str
}

guess_hfr_type <- function(x) {
  has_default <- (!is.null(x$header_default) || !is.null(x$footer_default))
  has_even <- (!is.null(x$header_even) || !is.null(x$footer_even))
  has_first <- (!is.null(x$header_first) || !is.null(x$footer_first))

  if (has_default && !has_even && !has_first) {
    "default"
  } else if (has_default && has_even && !has_first) {
    "evenodd"
  } else if (has_default && has_even && has_first) {
    "evenoddfirst"
  } else if (has_default && !has_even && has_first) {
    "defaultfirst"
  } else {
    "none"
  }
}

to_rtf.page_mar <- function(x, ...) {
  sprintf(
    "\\margb%.0f\\margt%.0f\\margr%.0f\\margl%.0f\\gutter%.0f\\headery%.0f\\footery%.0f",
    inch_to_tweep(x$bottom),
    inch_to_tweep(x$top),
    inch_to_tweep(x$right),
    inch_to_tweep(x$left),
    inch_to_tweep(x$gutter),
    inch_to_tweep(x$header),
    inch_to_tweep(x$footer)
  )
}

#' @export
to_rtf.page_size <- function(x, ...) {
  out <- sprintf(
    "\\pghsxn%.0f\\pgwsxn%.0f%s",
    inch_to_tweep(x$height),
    inch_to_tweep(x$width),
    if ("landscape" %in% x$orient) {
      "\\lndscpsxn"
    } else {
      ""
    }
  )
  out
}

#' @export
to_rtf.section_columns <- function(x, ...) {
  widths <- x$widths * 20 * 72
  space <- x$space * 20 * 72

  columns_str <- sprintf("\\colw%.0f\\colsx%.0f", widths[length(widths)], space)

  linebetcol <- ""
  if (x$sep) linebetcol <- "\\linebetcol"

  sprintf(
    "\\cols%.0f%s%s",
    length(widths),
    linebetcol,
    columns_str
  )
}

# RTF management ----
font_table <- function(families) {
  data.frame(
    family = families,
    index = seq_along(families) - 1,
    match = sprintf("%%font:%s%%", families),
    stringsAsFactors = FALSE
  )
}
color_table <- function(colors) {
  data.frame(
    color = colors,
    index = seq_along(colors),
    match_ftcolor = sprintf("%%ftcolor:%s%%", colors),
    match_ftshading = sprintf("%%ftshading:%s%%", colors),
    match_ftlinecolor = sprintf("%%ftlinecolor:%s%%", colors),
    match_ftbgcolor = sprintf("%%ftbgcolor:%s%%", colors),
    stringsAsFactors = FALSE
  )
}

ppr_rtf <- function(x) {
  text_align_ <- ""
  if ("left" %in% x$text.align) {
    text_align_ <- "\\ql"
  } else if ("center" %in% x$text.align) {
    text_align_ <- "\\qc"
  } else if ("right" %in% x$text.align) {
    text_align_ <- "\\qr"
  } else if ("justified" %in% x$text.align) {
    text_align_ <- "\\qj"
  }
  keep_with_next <- ""
  if (x$keep_with_next) {
    keep_with_next <- "\\keepn"
  }
  leftright_padding <- ""
  if (!is.na(x$padding.left) && !is.na(x$padding.right)) {
    leftright_padding <- sprintf(
      "\\fi0\\li%.0f\\ri%.0f",
      x$padding.left * 20, x$padding.right * 20
    )
  }
  topbot_spacing <- ""
  if (!is.na(x$padding.bottom) && !is.na(x$padding.top)) {
    topbot_spacing <- sprintf(
      "\\sb%.0f\\sa%.0f",
      x$padding.bottom * 20, x$padding.top * 20
    )
  }
  out <- paste0(
    sprintf("\\sl%.0f\\slmult1", 240 * x$line_spacing),
    text_align_,
    keep_with_next,
    topbot_spacing,
    leftright_padding
  )
  if (nchar(out) > 0) {
    out <- paste0(out, " ")
  }
  out
}

rpr_rtf <- function(x) {
  out <- ""

  if (!is.na(x$font.family)) {
    out <- paste0(out, "%font:", x$font.family, "%")
  }


  if (!is.na(x$italic)) {
    if (x$italic) {
      out <- paste0(out, "\\i")
    }
  }

  if (!is.na(x$bold)) {
    if (x$bold) {
      out <- paste0(out, "\\b")
    }
  }
  if (!is.na(x$underlined)) {
    if (x$underlined) {
      out <- paste0(out, "\\ul")
    }
  }

  if (x$vertical.align == "superscript") {
    out <- paste0(out, "\\super\\up", round(x$font.size / 2))
  } else if (x$vertical.align == "subscript") {
    out <- paste0(out, "\\sub\\dn", round(x$font.size / 2))
  }

  if (!is.na(x$font.size)) {
    out <- paste0(
      out,
      rtf_fontsize(x$font.size)
    )
  }

  if (!is.na(x$color)) {
    out <- paste0(
      out,
      paste0("%ftcolor:", x$color, "%")
    )
  }

  if (nchar(out) > 0) {
    out <- paste0(out, " ")
  }
  out
}

border_rtf <- function(x, side) {
  x$style[x$style %in% "solid"] <- "single"
  if (!x$style %in% c("dotted", "dashed", "single")) {
    x$style <- "single"
  }
  if (x$width < 0.0001 || is_transparent(x$color)) {
    return("")
  }

  linestyle <- c(
    "dotted" = "\\brdrdot",
    "dashed" = "\\brdrdash",
    "single" = "\\brdrs"
  )

  style_ <- linestyle[x$style]
  width_ <- paste0("\\brdrw", round(x$width * 20))
  color_ <- paste0("%ftlinecolor:", x$color, "%")

  paste0(
    "\\clbrdr", substr(side, 1, 1), style_,
    width_, color_
  )
}

tcpr_rtf <- function(x) {
  if (colalpha(x$background.color) > 0) {
    background.color <- paste0("%ftbgcolor:", x$background.color, "%")
  } else {
    background.color <- ""
  }
  vertalign <- c(
    "center" = "\\clvertalc",
    "top" = "\\clvertalt",
    "bottom" = "\\clvertalb"
  )
  vertical.align <- vertalign[x$vertical.align]

  text.direction <- ""
  text.direction[x$text.direction %in% "btlr"] <- "\\stextflow2"
  text.direction[x$text.direction %in% "tbrl"] <- "\\stextflow3"

  bb <- border_rtf(x$border.bottom, side = "bottom")
  bt <- border_rtf(x$border.top, side = "top")
  bl <- border_rtf(x$border.left, side = "left")
  br <- border_rtf(x$border.right, side = "right")

  rowspan <- ""
  if (x$rowspan > 1) {
    rowspan <- "\\clmgf"
  } else if (x$rowspan < 1) {
    rowspan <- "\\clmrg"
  }
  colspan <- ""
  if (x$colspan < 1) {
    colspan <- "\\clvmrg"
  } else {
    colspan <- "\\clvmgf"
  }

  paste0(
    bb, bt, bl, br,
    text.direction,
    rowspan, colspan,
    vertical.align, background.color
  )
}

# main ----
#' @export
#' @title Create an RTF document object
#' @description Creation of the object representing an
#' RTF document which can then receive contents with
#' the [rtf_add()] function and be written to a file with
#' the `print(x, target="doc.rtf")` function.
#' @param def_sec a [block_section] object used to defined default section.
#' @param normal_par an object generated by [fp_par()]
#' @param normal_chunk an object generated by [fp_text()]
#' @examples
#' rtf_doc(normal_par = fp_par(padding = 3))
#' @seealso [read_docx()], [print.rtf()], [rtf_add()]
#' @return an object of class `rtf` representing an
#' empty RTF document.
rtf_doc <- function(def_sec = prop_section(),
                    normal_par = fp_par(),
                    normal_chunk = fp_text(font.family = "Arial", font.size = 11)) {
  styles <- data.frame(
    style_name = "Normal",
    style_id = 1L,
    type = "paragraph",
    rtf = rtf_par_style(
      fp_p = normal_par,
      fp_t = normal_chunk
    ),
    stringsAsFactors = FALSE
  )

  x <- list(
    content = list(),
    default_section = def_sec,
    styles = styles,
    normal_par = normal_par,
    normal_chunk = normal_chunk
  )
  class(x) <- "rtf"
  x
}

#' @export
#' @title Add content into an RTF document
#' @description This function add 'officer' objects into an RTF document.
#' Values are added as new paragraphs. See section 'Methods (by class)'
#' that list supported objects.
#' @param x rtf object, created by [rtf_doc()].
#' @param value object to add in the document. Supported objects
#' are vectors, graphics, block of formatted paragraphs. Use package
#' 'flextable' to add tables.
#' @param ... further arguments passed to or from other methods. When
#' adding a `ggplot` object or [plot_instr], these arguments will be used
#' by png function. See section 'Methods' to see what arguments can be used.
#' @examples
#' library(officer)
#'
#' def_text <- fp_text_lite(color = "#006699", bold = TRUE)
#' center_par <- fp_par(text.align = "center", padding = 3)
#'
#' doc <- rtf_doc(
#'   normal_par = fp_par(line_spacing = 1.4, padding = 3)
#' )
#'
#' doc <- rtf_add(
#'   x = doc,
#'   value = fpar(
#'     ftext("how are you?", prop = def_text),
#'     fp_p = fp_par(text.align = "center")
#'   )
#' )
#'
#' a_paragraph <- fpar(
#'   ftext("Here is a date: ", prop = def_text),
#'   run_word_field(field = "Date \\@ \"MMMM d yyyy\""),
#'   fp_p = center_par
#' )
#' doc <- rtf_add(
#'   x = doc,
#'   value = block_list(
#'     a_paragraph,
#'     a_paragraph,
#'     a_paragraph
#'   )
#' )
#'
#' if (require("ggplot2")) {
#'   gg <- gg_plot <- ggplot(data = iris) +
#'     geom_point(mapping = aes(Sepal.Length, Petal.Length))
#'   doc <- rtf_add(doc, gg,
#'     width = 3, height = 4,
#'     ppr = center_par
#'   )
#' }
#' anyplot <- plot_instr(code = {
#'   barplot(1:5, col = 2:6)
#' })
#' doc <- rtf_add(doc, anyplot,
#'   width = 5, height = 4,
#'   ppr = center_par
#' )
#'
#' print(doc, target = tempfile(fileext = ".rtf"))
rtf_add <- function(x, value, ...) {
  UseMethod("rtf_add", value)
}


#' @export
#' @describeIn rtf_add add a new section definition
rtf_add.block_section <- function(x, value, ...) {
  x$content[[length(x$content) + 1]] <- value
  x
}

#' @export
#' @describeIn rtf_add add characters as new paragraphs
rtf_add.character <- function(x, value, ...) {
  values <- lapply(
    X = str_encode_to_rtf(value),
    function(z) {
      fpar(ftext(z, prop = x$normal_chunk), fp_p = x$normal_par)
    }
  )

  x$content <- append(x$content, values)
  x
}

#' @export
#' @describeIn rtf_add add a factor vector as new paragraphs
rtf_add.factor <- function(x, value, ...) {
  values <- lapply(
    X = str_encode_to_rtf(as.character(value)),
    function(z) {
      fpar(ftext(z, prop = x$normal_chunk), fp_p = x$normal_par)
    }
  )

  x$content <- append(x$content, values)
  x
}

#' @export
#' @describeIn rtf_add add a double vector as new paragraphs
#' @param formatter function used to format the numerical values
rtf_add.double <- function(x, value, formatter = formatC, ...) {
  values <- lapply(
    X = value,
    function(z) {
      fpar(ftext(formatC(z), prop = x$normal_chunk), fp_p = x$normal_par)
    }
  )

  x$content <- append(x$content, values)
  x
}

#' @export
#' @describeIn rtf_add add an [fpar()]
rtf_add.fpar <- function(x, value, ...) {
  x$content[[length(x$content) + 1]] <- value
  x
}

#' @export
#' @describeIn rtf_add add an [block_list()]
rtf_add.block_list <- function(x, value, ...) {
  x$content <- append(x$content, as.list(value))
  x
}


#' @export
#' @describeIn rtf_add add a ggplot2
#' @param width height in inches
#' @param height height in inches
#' @param res resolution of the png image in ppi
#' @param scale Multiplicative scaling factor, same as in ggsave
#' @param ppr [fp_par()] to apply to paragraph.
rtf_add.gg <- function(x, value, width = 6, height = 5, res = 300, scale = 1, ppr = fp_par(text.align = "center"), ...) {
  if (!requireNamespace("ggplot2")) {
    stop("package ggplot2 is required to use this function")
  }

  file <- tempfile(fileext = ".png")
  agg_png(filename = file, width = width, height = height, units = "in", res = res, scaling = scale, background = "transparent", ...)
  tryCatch(
    {
      print(value)
    },
    finally = {
      dev.off()
    }
  )

  value <- fpar(
    external_img(src = file, width = width, height = height, guess_size = FALSE),
    fp_p = ppr
  )

  x$content[[length(x$content) + 1]] <- value
  x
}
#' @export
#' @describeIn rtf_add add a [plot_instr()] object
rtf_add.plot_instr <- function(x, value, width = 6, height = 5, res = 300, scale = 1, ppr = fp_par(text.align = "center"), ...) {
  file <- tempfile(fileext = ".png")
  agg_png(filename = file, width = width, height = height, units = "in", res = res, scaling = scale, background = "transparent", ...)
  tryCatch(
    {
      eval(value$code)
    },
    finally = {
      dev.off()
    }
  )

  value <- fpar(
    external_img(src = file, width = width, height = height, guess_size = FALSE),
    fp_p = ppr
  )

  x$content[[length(x$content) + 1]] <- value
  x
}

#' @title Write an 'RTF' document to a file
#' @description Write the RTF object and its content to a file.
#' @param x an 'rtf' object created with [rtf_doc()]
#' @param target path to the RTF file to write
#' @param ... unused
#' @examples
#' # write a rdocx object in a rtf file ----
#' doc <- rtf_doc()
#' print(doc, target = tempfile(fileext = ".rtf"))
#' @export
#' @seealso [rtf_doc()]
print.rtf <- function(x, target = NULL, ...) {
  if (is.null(target)) {
    cat("rtf document with", length(x$content), "element(s)\n")
    return(invisible())
  }

  rtf_content <- sapply(x$content, to_rtf)
  rtf_ss <- rtf_stylesheet(x)
  txt_to_scan <- c(rtf_content, rtf_ss)
  m <- gregexec(pattern = "%font:[a-zA-Z ]+%", text = txt_to_scan)
  fonts <- unique(unlist(regmatches(txt_to_scan, m)))
  fonts <- gsub("(^%font:|%$)", "", fonts)
  family_table <- font_table(families = fonts)

  m <- gregexec(pattern = "\\%ftcolor\\:[a-zA-Z #0-9]+\\%", text = txt_to_scan)
  ftcolor <- unlist(regmatches(txt_to_scan, m))
  m <- gregexec(pattern = "\\%ftshading\\:[a-zA-Z #0-9]+\\%", text = txt_to_scan)
  ftshading <- unlist(regmatches(txt_to_scan, m))
  m <- gregexec(pattern = "\\%ftlinecolor\\:[a-zA-Z #0-9]+\\%", text = txt_to_scan)
  ftlinecolor <- unlist(regmatches(txt_to_scan, m))
  m <- gregexec(pattern = "\\%ftbgcolor\\:[a-zA-Z #0-9]+\\%", text = txt_to_scan)
  ftbgcolor <- unlist(regmatches(txt_to_scan, m))

  supported_highlight <- c(
    "black", "blue", "cyan", "green", "magenta", "red",
    "yellow", "darkblue", "darkcyan", "darkgreen",
    "darkmagenta", "darkred", "#8B8000", "darkgrey", "lightgray"
  )

  colors <- unique(c(ftcolor, ftshading, ftlinecolor, ftbgcolor, supported_highlight))
  colors <- gsub("(^%(ftcolor|ftshading|ftlinecolor|ftbgcolor):|%$)", "", colors)
  color_table <- color_table(colors = colors)

  header <- c(
    "{\\rtf1\\deff0",
    rtfize_font_table(family_table),
    rtfize_color_table(color_table)
  )

  default_sect <- to_rtf(x$default_section)
  all_strings <- c(
    header, rtf_ss, default_sect, rtf_content,
    attr(default_sect, "hfr"), "}"
  )

  all_strings <- fix_font_ref(all_strings, family_table)
  all_strings <- fix_font_color(all_strings, color_table)
  all_strings <- fix_line_color(all_strings, color_table)
  all_strings <- fix_bg_color(all_strings, color_table)
  all_strings <- fix_font_shading(all_strings, color_table)

  writeLines(all_strings, target, useBytes = TRUE)
  target
}

#' @export
#' @title Encode UTF8 string to RTF
#' @description Convert strings to RTF valid codes.
#' @param z character vector to be converted
#' @family functions for officer extensions
#' @return character vector of results encoded to RTF
#' @keywords internal
#' @examples
#' str_encode_to_rtf("Hello World")
str_encode_to_rtf <- function(z) {

  # https://stackoverflow.com/questions/1368020/how-to-output-unicode-string-to-rtf-using-c
  char_list <- strsplit(z, "")
  int_list <- lapply(z, utf8ToInt)
  rtf_strings <- mapply(
    function(charv, intv) {
      is_tobeprotected <- charv %in% c("\\", "{", "}")
      # is_std_ansi <- intv < 128
      # is_ext_ansi <- !is_std_ansi & intv < 256
      is_ext_ansi <- intv < 256
      is_unic2b <- !is_ext_ansi & intv < 32768
      is_unic4b <- !is_ext_ansi & !is_unic2b
      charv[is_unic2b] <- paste0("\\uc1\\u", intv[is_unic2b], "?")
      charv[is_unic4b] <- paste0("\\uc1\\u", intv[is_unic4b] - 65536, "?")
      paste0(charv, collapse = "")
    }, char_list, int_list,
    SIMPLIFY = TRUE
  )
  as.character(rtf_strings)
}


# tools ----
rtf_fontsize <- function(x) {
  size <- formatC(x * 2, format = "f", digits = 0)
  z <- sprintf("\\fs%s", size)
  z[is.na(x)] <- ""
  z
}
rtf_color <- function(x, color_table, mark = "cf") {
  index <- color_table$index[match(x, color_table$color)]
  z <- sprintf("\\%s%.0f", mark, index)
  z[is.na(index)] <- ""
  z
}

rtfize_font_table <- function(x) {
  z <- sprintf("{\\f%.0f %s;}", x$index, x$family)
  z <- paste0(z, collapse = "")
  paste0("{\\fonttbl", z, "}")
}

rtf_color_code <- function(color) {
  color <- as.vector(col2rgb(color, alpha = FALSE))
  sprintf(
    "\\red%.0f\\green%.0f\\blue%.0f;",
    color[1], color[2], color[3]
  )
}

rtfize_color_table <- function(x) {
  z <- vapply(x$color, rtf_color_code, "")
  z <- paste0(z, collapse = "")
  paste0("{\\colortbl;", z, "}")
}

fix_font_ref <- function(x, font_tbl) {
  for (i in seq_len(nrow(font_tbl))) {
    m <- gregexec(pattern = font_tbl$match[i], text = x)
    regmatches(x, m) <- paste0("\\f", font_tbl$index[i])
  }
  x
}
fix_font_color <- function(x, color_tbl) {
  for (i in seq_len(nrow(color_tbl))) {
    m <- gregexec(pattern = color_tbl$match_ftcolor[i], text = x)
    regmatches(x, m) <- paste0("\\cf", color_tbl$index[i])
  }
  x
}
fix_line_color <- function(x, color_tbl) {
  for (i in seq_len(nrow(color_tbl))) {
    m <- gregexec(pattern = color_tbl$match_ftlinecolor[i], text = x)
    regmatches(x, m) <- paste0("\\brdrcf", color_tbl$index[i])
  }
  x
}
fix_bg_color <- function(x, color_tbl) {
  for (i in seq_len(nrow(color_tbl))) {
    m <- gregexec(pattern = color_tbl$match_ftbgcolor[i], text = x)
    regmatches(x, m) <- paste0("\\clcbpat", color_tbl$index[i])
  }
  x
}
fix_font_shading <- function(x, color_tbl) {
  supported_highlight <- c(
    "black", "blue", "cyan", "green", "magenta", "red",
    "yellow", "darkblue", "darkcyan", "darkgreen",
    "darkmagenta", "darkred", "#8B8000", "darkgrey", "lightgray"
  )

  supported_highlight <-
    data.frame(
      colors = supported_highlight,
      index = seq_along(supported_highlight),
      match_ftshading = sprintf("%%ftshading:%s%%", supported_highlight)
    )

  for (i in seq_len(nrow(color_tbl))) {
    m <- gregexec(pattern = color_tbl$match_ftshading[i], text = x)
    matches <- Filter(function(x) x > -1, m)
    if (length(matches) > 0) {
      used_shading <- unique(unlist(regmatches(x, m)))
      if (!used_shading %in% c(supported_highlight$match_ftshading, "%ftshading:transparent%")) {
        stop("chunk highlight color (", used_shading, ") in RTF can only have values: ", paste0(shQuote(supported_highlight$colors), collapse = ", "))
      }
    }

    regmatches(x, m) <- sprintf("\\highlight%.0f", color_tbl$index[i])
  }
  x <- gsub("%ftshading:transparent%", "", x, fixed = TRUE)
  x
}

rtf_par_style <- function(fp_p = fp_par(), fp_t = NULL) {
  if (!is.null(fp_t)) {
    fp_t_rtf <- rpr_rtf(fp_t)
  } else {
    fp_t_rtf <- ""
  }

  paste0(ppr_rtf(fp_p), fp_t_rtf)
}

rtf_set_paragraph_style <- function(x, style_name, fp_p = fp_par(), fp_t = NULL) {
  index <- which(x$styles$style_name %in% style_name)
  style_id <- if (length(index) < 1) {
    style_id <- nrow(x$styles) + 1L
  } else {
    style_id <- index
  }
  new_style <- data.frame(
    style_name = style_name,
    style_id = style_id,
    type = "paragraph",
    rtf = rtf_par_style(fp_p = fp_p, fp_t = fp_t),
    stringsAsFactors = FALSE
  )

  if (length(index) < 1) {
    x$styles <- rbind(x$styles, new_style)
  } else {
    x$styles[index, "rtf"] <- new_style$rtf
  }

  x
}

rtf_stylesheet <- function(x) {
  out <- character(nrow(x$styles))
  for (i in seq_len(nrow(x$styles))) {
    out[i] <- sprintf(
      "\n{\\f%.0f\\s%.0f%s %s;}",
      i, i, x$styles$rtf[i],
      x$styles$style_name[i]
    )
  }
  paste0(
    "{\\stylesheet ",
    paste0(out, collapse = ""),
    "}"
  )
}


# check for gregexec -----
if (!"gregexec" %in% getNamespaceExports("base")) {
  # copied from R source, grep.R
  gregexec <- function(pattern, text, ignore.case = FALSE, perl = FALSE,
                       fixed = FALSE, useBytes = FALSE) {
    if(is.factor(text) && length(levels(text)) < length(text)) {
      out <- gregexec(pattern, c(levels(text), NA_character_),
                      ignore.case, perl, fixed, useBytes)
      outna <- out[length(out)]
      out <- out[text]
      out[is.na(text)] <- outna
      return(out)
    }

    dat <- gregexpr(pattern = pattern, text=text, ignore.case = ignore.case,
                    fixed = fixed, useBytes = useBytes, perl = perl)
    if(perl && !fixed) {
      ## Perl generates match data, so use that
      capt.attr <- c('capture.start', 'capture.length', 'capture.names')
      process <- function(x) {
        if(anyNA(x) || any(x < 0)) y <- x
        else {
          ## Interleave matches with captures
          y <- t(cbind(x, attr(x, "capture.start")))
          attributes(y)[names(attributes(x))] <- attributes(x)
          ml <- t(cbind(attr(x, "match.length"), attr(x, "capture.length")))
          nm <- attr(x, 'capture.names')
          ## Remove empty names that `gregexpr` returns
          dimnames(ml) <- dimnames(y) <-
            if(any(nzchar(nm))) list(c("", nm), NULL)
          attr(y, "match.length") <- ml
          y
        }
        attributes(y)[capt.attr] <- NULL
        y
      }
      lapply(dat, process)
    } else {
      ## For TRE or fixed we must compute the match data ourselves
      m1 <- lapply(regmatches(text, dat),
                   regexec, pattern = pattern, ignore.case = ignore.case,
                   perl = perl, fixed = fixed, useBytes = useBytes)
      mlen <- lengths(m1)
      res <- vector("list", length(m1))
      im <- mlen > 0
      res[!im] <- dat[!im]   # -1, NA
      res[im] <- Map(
        function(outer, inner) {
          tmp <- do.call(cbind, inner)
          attributes(tmp)[names(attributes(inner))] <- attributes(inner)
          attr(tmp, 'match.length') <-
            do.call(cbind, lapply(inner, `attr`, 'match.length'))
          # useBytes/index.type should be same for all so use outer vals
          attr(tmp, 'useBytes') <- attr(outer, 'useBytes')
          attr(tmp, 'index.type') <- attr(outer, 'index.type')
          tmp + rep(outer - 1L, each = nrow(tmp))
        },
        dat[im],
        m1[im]
      )
      res
    }
  }
}
