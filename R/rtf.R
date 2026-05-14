# to_rtf ----
#' @export
#' @title Convert officer objects to RTF
#' @description Convert an object made with package officer
#' to RTF.
#' @param x object to convert to RTF. Supported
#' objects are:
#' - [ftext()]
#' - [external_img()]
#' - [run_autonum()]
#' - [run_columnbreak()]
#' - [run_linebreak()]
#' - [run_word_field()]
#' - [run_reference()]
#' - [run_pagebreak()]
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
  style_index <- attr(x, "rtf_style_index")
  if (!is.null(style_index)) {
    # Canonical RTF emission for a styled paragraph (RTF spec, additive
    # model): \pard\plain \sN <style properties duplicated> text \par.
    # The duplicated properties make readers that ignore styles still
    # render the formatting, and prevent Word from recording spurious
    # "overrides" against the named style.
    style_props <- attr(x, "rtf_style_props")
    if (is.null(style_props)) style_props <- ""
    return(paste0(
      sprintf("{\\pard\\plain\\s%.0f", style_index),
      style_props,
      " ",
      rtf_chunks,
      "\\par}"
    ))
  }
  paste0("{", ppr_rtf(x$fp_p), rtf_chunks, "\\par}")
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
to_rtf.run_tab <- function(x, ...) {
  "\\tab "
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
    width * 1440,
    height * 1440
  )
  paste(rtf_str, dat, "\n}", sep = "")
}


#' @export
to_rtf.floating_external_img <- function(x, ...) {
  dims <- attr(x, "dims")
  width <- dims$width
  height <- dims$height

  pos <- attr(x, "pos")
  pos_x <- pos$x
  pos_y <- pos$y
  pos_h_from <- pos$h_from
  pos_v_from <- pos$v_from

  wrap <- attr(x, "wrap")
  wrap_type <- wrap$type
  wrap_side <- wrap$side
  wrap_dist_top <- wrap$dist_top
  wrap_dist_bottom <- wrap$dist_bottom
  wrap_dist_left <- wrap$dist_left
  wrap_dist_right <- wrap$dist_right

  if (!grepl("\\.png$", x)) {
    imgpath <- tempfile(fileext = ".png")
    img <- magick::image_read(x)
    magick::image_write(img, path = imgpath)
  } else {
    imgpath <- x
  }

  # Read image data
  dat <- readBin(imgpath, what = "raw", size = 1, endian = "little", n = 1e+8)
  dat <- paste(dat, collapse = "")

  # RTF units: twips (1/1440 inch)
  width_twips <- width * 1440
  height_twips <- height * 1440
  pos_x_twips <- pos_x * 1440
  pos_y_twips <- pos_y * 1440

  # EMU units for wrap distances (1 EMU = 1/914400 inch, so multiply inches by 914400)
  # But for RTF {\sp{\sn property}{\sv value}}, wrap distances are in EMUs
  dist_t_emus <- wrap_dist_top * 914400
  dist_b_emus <- wrap_dist_bottom * 914400
  dist_l_emus <- wrap_dist_left * 914400
  dist_r_emus <- wrap_dist_right * 914400

  # RTF wrap codes for \shpwr:
  # 1 = wrap around top and bottom
  # 2 = wrap around shape
  # 3 = none (as if shape isn't present)
  # 4 = wrap tightly
  # 5 = wrap text through shape
  wrap_code <- switch(
    wrap_type,
    "square" = "\\shpwr2",
    "tight" = "\\shpwr4",
    "topAndBottom" = "\\shpwr1",
    "through" = "\\shpwr5",
    "none" = "\\shpwr3",
    "\\shpwr2" # Default to square wrap
  )

  # RTF wrap side codes for \shpwrk:
  # 0 = wrap both sides
  # 1 = wrap left side only
  # 2 = wrap right side only
  # 3 = wrap only on largest side
  wrap_side_code <- switch(
    wrap_side,
    "bothSides" = "\\shpwrk0",
    "left" = "\\shpwrk1",
    "right" = "\\shpwrk2",
    "largest" = "\\shpwrk3",
    "\\shpwrk0" # Default
  )

  # RTF positioning codes
  # Horizontal positioning relative to:
  # \shpbxpage = page, \shpbxmargin = margin, \shpbxcolumn = column
  pos_h_code <- switch(
    pos_h_from,
    "page" = "\\shpbxpage",
    "margin" = "\\shpbxmargin",
    "column" = "\\shpbxcolumn",
    "character" = "\\shpbxcolumn",
    "\\shpbxmargin" # Default
  )

  # Vertical positioning relative to:
  # \shpbypage = page, \shpbymargin = margin, \shpbypara = paragraph
  pos_v_code <- switch(
    pos_v_from,
    "page" = "\\shpbypage",
    "margin" = "\\shpbymargin",
    "paragraph" = "\\shpbypara",
    "line" = "\\shpbypara",
    "\\shpbymargin" # Default
  )

  # posrelh property: 0=margin, 1=page, 2=column
  posrelh <- switch(
    pos_h_from,
    "margin" = 0,
    "page" = 1,
    "column" = 2,
    "character" = 2,
    0 # Default to margin
  )

  # posrelv property: 0=margin, 1=page, 2=paragraph
  posrelv <- switch(
    pos_v_from,
    "margin" = 0,
    "page" = 1,
    "paragraph" = 2,
    "line" = 2,
    0 # Default to margin
  )

  # Build the RTF shape structure following Word's format:
  # {\shp \shpleft... \shptop... \shpright... \shpbottom... \shpfhdr0
  #       \shpbx... \shpbxignore \shpby... \shpbyignore
  #       \shpwr... \shpwrk... \shpfblwtxt... \shpz... \shplid...
  #       {\*\shpinst {\sp{\sn prop}{\sv val}}... \par}}}

  rtf_str <- paste0(
    "{\\shp",
    sprintf(
      "\\shpleft%.0f\\shptop%.0f\\shpright%.0f\\shpbottom%.0f",
      pos_x_twips,
      pos_y_twips,
      pos_x_twips + width_twips,
      pos_y_twips + height_twips
    ),
    "\\shpfhdr0", # Not in header/footer
    pos_h_code,
    "\\shpbxignore", # Horizontal positioning
    pos_v_code,
    "\\shpbyignore", # Vertical positioning
    wrap_code,
    wrap_side_code, # Wrap type and side
    "\\shpfblwtxt0", # Image in front of text
    "\\shpz0", # Z-order
    "\\shplid1027" # Shape ID
  )

  # Start the shape properties ({\*\shpinst ...})
  rtf_str <- paste0(rtf_str, "{\\*\\shpinst")

  # Add shape properties following {\sp{\sn PropertyName}{\sv PropertyValue}} format
  # shapeType: 75 = picture frame
  rtf_str <- paste0(rtf_str, "{\\sp{\\sn shapeType}{\\sv 75}}")

  # pib: Binary picture data
  rtf_str <- paste0(
    rtf_str,
    sprintf(
      "{\\sp{\\sn pib}{\\sv {\\pict\\pngblip\\picwgoal%.0f\\pichgoal%.0f %s}}}",
      width_twips,
      height_twips,
      dat
    )
  )

  # Add wrap distances (in EMUs)
  if (dist_l_emus > 0) {
    rtf_str <- paste0(
      rtf_str,
      sprintf("{\\sp{\\sn dxWrapDistLeft}{\\sv %.0f}}", dist_l_emus)
    )
  }
  if (dist_r_emus > 0) {
    rtf_str <- paste0(
      rtf_str,
      sprintf("{\\sp{\\sn dxWrapDistRight}{\\sv %.0f}}", dist_r_emus)
    )
  }
  if (dist_t_emus > 0) {
    rtf_str <- paste0(
      rtf_str,
      sprintf("{\\sp{\\sn dyWrapDistTop}{\\sv %.0f}}", dist_t_emus)
    )
  }
  if (dist_b_emus > 0) {
    rtf_str <- paste0(
      rtf_str,
      sprintf("{\\sp{\\sn dyWrapDistBottom}{\\sv %.0f}}", dist_b_emus)
    )
  }

  # Add positioning properties (posrelh, posrelv are not in the Word example,
  # but keeping them won't hurt for compatibility with newer readers)
  # Note: These are commented out since the Word example doesn't include them
  # rtf_str <- paste0(rtf_str, sprintf("{\\sp{\\sn posrelh}{\\sv %d}}", posrelh))
  # rtf_str <- paste0(rtf_str, sprintf("{\\sp{\\sn posrelv}{\\sv %d}}", posrelv))

  # fBehindDocument: 0 = in front of text, 1 = behind text
  # Note: This is redundant with \shpfblwtxt0 above, but Word includes it
  # rtf_str <- paste0(rtf_str, "{\\sp{\\sn fBehindDocument}{\\sv 0}}")

  # Close the shape structure: \par}}} closes {\*\shpinst and {\shp
  paste0(rtf_str, "\\par}}}")
}


#' @export
to_rtf.block_toc <- function(x, ...) {
  # Field switches must be written with doubled backslashes inside
  # \fldinst (e.g. \\o, \\h, \\z, \\u). Single-backslash forms collide
  # with RTF control words (\o = overprint, \h = hidden, \u = Unicode
  # character), and Word will reject the document as malformed.
  field <- if (is.null(x$style) && is.null(x$seq_id)) {
    sprintf("TOC \\\\o \"1-%.0f\" \\\\h \\\\z \\\\u", x$level)
  } else if (!is.null(x$style)) {
    sprintf("TOC \\\\h \\\\z \\\\t \"%s\"", x$style)
  } else {
    sprintf("TOC \\\\h \\\\z \\\\c \"%s\"", x$seq_id)
  }
  paste0("{\\pard ", to_rtf(run_word_field(field = field)), "\\par}")
}

#' @export
to_rtf.run_word_field <- function(x, ...) {
  rtf_fp <- ""
  if (!is.null(x$pr)) {
    rtf_fp <- rpr_rtf(x$pr)
  }
  # The \fldrslt group is required by the RTF spec for a field group:
  # {\field {\*\fldinst <instructions>} {\fldrslt <result>}}. Word
  # tolerates its absence for simple fields like PAGE but rejects the
  # document for complex fields like TOC. An empty result is valid;
  # Word recalculates it when the field is updated.
  sprintf("%s{\\field{\\*\\fldinst %s}{\\fldrslt }}", rtf_fp, x$field)
}

#' @export
to_rtf.run_pagebreak <- function(x, ...) {
  "\\page"
}
#' @export
to_rtf.run_columnbreak <- function(x, ...) {
  "\\column "
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
  if (is.null(x$seq_id)) {
    return("")
  }

  run_str_pre <- to_rtf(ftext(x$pre_label, prop = x$pr))
  run_str_post <- to_rtf(ftext(x$post_label, prop = x$pr))

  seq_str <- paste0("SEQ ", x$seq_id, " \\\\* Arabic")
  if (!is.null(x$start_at) && is.numeric(x$start_at)) {
    seq_str <- paste(seq_str, "\\r", as.integer(x$start_at))
  }

  sqf <- run_word_field(field = seq_str, prop = x$pr)
  sf_str <- to_rtf(sqf)

  if (x$tnd > 0) {
    z <- paste0(
      to_rtf(run_word_field(
        field = paste0("STYLEREF ", x$tnd, " \\r"),
        prop = x$pr
      )),
      to_rtf(ftext(x$tns, prop = x$pr))
    )
    sf_str <- paste0(z, sf_str)
  }

  if (!is.null(x$bookmark)) {
    if (x$bookmark_all) {
      out <- paste0(
        rtf_bookmark(id = x$bookmark, str = paste0(run_str_pre, sf_str)),
        run_str_post
      )
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
  out <- to_rtf(run_word_field(
    field = paste0(" REF ", x$id, " \\\\h "),
    prop = x$pr
  ))
  out
}

## to_rtf sections ----

#' @export
to_rtf.block_section <- function(x, ...) {
  # API convention: block_section() configures the section that
  # *follows* (content added after it). In RTF that maps to
  # \sect (close previous section) then the property keywords,
  # which apply to the newly opened section. \sectd resets section
  # properties to defaults so that a prop_section() without columns
  # does not inherit \cols2 from a previous column section.
  paste0("\\sect\\sectd", to_rtf(x$property))
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
  str_h <- ""
  str_f <- ""
  str_h_even <- ""
  str_f_even <- ""
  str_h_first <- ""
  str_f_first <- ""

  if ("none" %in% hfr_type) {
    hfr_str <- ""
  } else if ("default" %in% hfr_type) {
    if (!is.null(x$header_default)) {
      str_h <- paste0("{\\header ", to_rtf(x$header_default), "}")
    }
    if (!is.null(x$footer_default)) {
      str_f <- paste0("{\\footer ", to_rtf(x$footer_default), "}")
    }
    hfr_str <- paste0(str_h, str_f)
  } else if ("evenodd" %in% hfr_type) {
    extra_rtf <- "\\facingp"
    if (!is.null(x$header_default)) {
      str_h <- paste0("{\\headerl ", to_rtf(x$header_default), "}")
    }
    if (!is.null(x$header_even)) {
      str_h_even <- paste0("{\\headerr ", to_rtf(x$header_even), "}")
    }
    if (!is.null(x$footer_default)) {
      str_f <- paste0("{\\footerl ", to_rtf(x$footer_default), "}")
    }
    if (!is.null(x$footer_even)) {
      str_f_even <- paste0("{\\footerr ", to_rtf(x$footer_even), "}")
    }
    hfr_str <- paste0(
      str_h,
      str_h_even,
      str_f,
      str_f_even
    )
  } else if ("evenoddfirst" %in% hfr_type) {
    extra_rtf <- "\\facingp\\titlepg"
    if (!is.null(x$header_default)) {
      str_h <- paste0("{\\headerl ", to_rtf(x$header_default), "}")
    }
    if (!is.null(x$header_even)) {
      str_h_even <- paste0("{\\headerr ", to_rtf(x$header_even), "}")
    }
    if (!is.null(x$header_first)) {
      str_h_first <- paste0("{\\headerf ", to_rtf(x$header_first), "}")
    }
    if (!is.null(x$footer_default)) {
      str_f <- paste0("{\\footerl ", to_rtf(x$footer_default), "}")
    }
    if (!is.null(x$footer_even)) {
      str_f_even <- paste0("{\\footerr ", to_rtf(x$footer_even), "}")
    }
    if (!is.null(x$footer_first)) {
      str_f_first <- paste0("{\\footerf ", to_rtf(x$footer_first), "}")
    }
    hfr_str <- paste0(
      str_h,
      str_h_even,
      str_h_first,
      str_f,
      str_f_even,
      str_f_first
    )
  } else {
    extra_rtf <- "\\titlepg"
    if (!is.null(x$header_default)) {
      str_h <- paste0("{\\headerl ", to_rtf(x$header_default), "}")
    }
    if (!is.null(x$header_first)) {
      str_h_first <- paste0("{\\headerf ", to_rtf(x$header_first), "}")
    }
    if (!is.null(x$footer_default)) {
      str_f <- paste0("{\\footerl ", to_rtf(x$footer_default), "}")
    }
    if (!is.null(x$footer_first)) {
      str_f_first <- paste0("{\\footerf ", to_rtf(x$footer_first), "}")
    }
    hfr_str <- paste0(
      str_h,
      str_h_first,
      str_f,
      str_f_first
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

#' @export
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
  if (!is.null(x$width) && !is.null(x$height)) {
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
  } else {
    out <- if ("landscape" %in% x$orient) {
      "\\lndscpsxn"
    } else {
      ""
    }
  }

  out
}

#' @export
to_rtf.section_columns <- function(x, ...) {
  widths <- x$widths * 20 * 72
  space <- x$space * 20 * 72

  # RTF needs one <\colno N \colw N> per column for unequal widths.
  # \colsr (per-column right space) duplicates the global \colsx when
  # spaces are uniform, but matches what Word/LibreOffice emit and
  # leaves room for per-column spacing later. The last column has no
  # trailing \colsr.
  columns_str_all_but_last <- sprintf(
    "\\colno%.0f\\colw%.0f\\colsr%.0f",
    seq(length(widths) - 1),
    widths[-length(widths)],
    rep(space, length(widths) - 1)
  )
  columns_str_last <- sprintf(
    "\\colno%.0f\\colw%.0f",
    length(widths),
    widths[length(widths)]
  )
  columns_str <- c(columns_str_all_but_last, columns_str_last)

  linebetcol <- ""
  if (x$sep) {
    linebetcol <- "\\linebetcol"
  }

  sprintf(
    "\\cols%.0f\\colsx%.0f%s%s",
    length(widths),
    space,
    linebetcol,
    paste(columns_str, collapse = "")
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

ppr_rtf <- function(x, include_pard = TRUE) {
  text_align_ <- ""
  if (!is.na(x$text.align)) {
    if ("left" %in% x$text.align) {
      text_align_ <- "\\ql"
    } else if ("center" %in% x$text.align) {
      text_align_ <- "\\qc"
    } else if ("right" %in% x$text.align) {
      text_align_ <- "\\qr"
    } else if ("justified" %in% x$text.align) {
      text_align_ <- "\\qj"
    }
  }

  tabs <- ""
  if (!is.null(x$tabs) && length(x$tabs) > 0 && !isFALSE(x$tabs)) {
    tabs <- paste0(sapply(x$tabs, rtf_fp_tab), collapse = "")
  }

  keep_with_next <- ""
  if (isTRUE(x$keep_with_next)) {
    keep_with_next <- "\\keepn"
  }
  leftright_padding <- ""
  if (!is.na(x$padding.left) && !is.na(x$padding.right)) {
    hang <- x$hanging %||% NA_real_
    first <- x$first_line %||% NA_real_
    if (!is.na(hang) && hang > 0) {
      fi <- sprintf("\\fi-%.0f", hang * 20)
    } else if (!is.na(first) && first > 0) {
      fi <- sprintf("\\fi%.0f", first * 20)
    } else {
      fi <- "\\fi0"
    }
    leftright_padding <- sprintf(
      "%s\\li%.0f\\ri%.0f",
      fi,
      x$padding.left * 20,
      x$padding.right * 20
    )
  }
  topbot_spacing <- ""
  if (!is.na(x$padding.bottom) && !is.na(x$padding.top)) {
    topbot_spacing <- sprintf(
      "\\sb%.0f\\sa%.0f",
      x$padding.bottom * 20,
      x$padding.top * 20
    )
  }

  line_spacing <- ""
  if (!is.na(x$line_spacing)) {
    line_spacing <- sprintf("\\sl%.0f\\slmult1", 240 * x$line_spacing)
  }

  out <- paste0(
    line_spacing,
    text_align_,
    tabs,
    keep_with_next,
    topbot_spacing,
    leftright_padding
  )
  if (nchar(out) > 0) {
    prefix <- if (isTRUE(include_pard)) "\\pard" else ""
    out <- paste0(prefix, out, " ")
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
  if (!is.na(x$strike)) {
    if (x$strike) {
      out <- paste0(out, "\\strike")
    }
  }

  if (x$vertical.align == "superscript" && !is.na(x$font.size)) {
    out <- paste0(out, sprintf("\\super\\up%.0f", round(x$font.size / 2)))
  } else if (x$vertical.align == "subscript" && !is.na(x$font.size)) {
    out <- paste0(out, sprintf("\\sub\\dn%.0f", round(x$font.size / 2)))
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
    "\\clbrdr",
    substr(side, 1, 1),
    style_,
    width_,
    color_
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
    bb,
    bt,
    bl,
    br,
    text.direction,
    rowspan,
    colspan,
    vertical.align,
    background.color
  )
}

# main ----
#' @export
#' @title Create an RTF document object
#' @description Creation of the object representing an
#' RTF document which can then receive contents with
#' the [rtf_add()] function and be written to a file with
#' the `print(x, target="doc.rtf")` function.
#' @param def_sec a [prop_section] object used to define the default section
#' (page size, margins, header / footer, columns) applied to content added
#' before any explicit [block_section()].
#' @param normal_par an object generated by [fp_par()]
#' @param normal_chunk an object generated by [fp_text()]
#'
#' @section Built-in styles:
#' `rtf_doc()` registers four paragraph styles: `"Normal"` (built from
#' `normal_par` / `normal_chunk`) and `"heading 1"` / `"heading 2"` /
#' `"heading 3"` (bold, dark blue shades, with outline levels so the
#' paragraphs feed Word's navigation pane and TOC field). The heading
#' names mirror Word's own convention so a script that adds paragraphs
#' with `style = "heading 1"` works identically against Word output
#' (`body_add_par(..., style = "heading 1")`) and RTF output
#' (`rtf_add(..., style = "heading 1")`).
#'
#' To override a built-in look (size, color, spacing, outline level),
#' call [rtf_set_paragraph_style()] after `rtf_doc()` with the same
#' `style_id`.
#'
#' ```r
#' doc <- rtf_doc()
#' doc <- rtf_set_paragraph_style(
#'   doc,
#'   style_id = "heading 1",
#'   fp_p = fp_par(text.align = "center", padding = 12),
#'   fp_t = fp_text_lite(bold = TRUE, font.size = 24, color = "#000000"),
#'   outline_level = 1L
#' )
#' ```
#'
#' Use [rtf_styles_info()] to inspect the current style table.
#' @examples
#' rtf_doc(normal_par = fp_par(padding = 3))
#' @seealso [read_docx()], [print.rtf()], [rtf_add()]. See `?rtf_add` for a
#' complete multi-section example exercising `def_sec`, headers / footers
#' and several `block_section()` calls.
#' @return an object of class `rtf` representing an
#' empty RTF document.
rtf_doc <- function(
  def_sec = prop_section(),
  normal_par = fp_par(),
  normal_chunk = fp_text(font.family = "Arial", font.size = 11)
) {
  styles <- data.frame(
    style_id = "Normal",
    style_name = "Normal",
    type = "paragraph",
    rtf_index = 1L,
    based_on = NA_character_,
    rtf = rtf_par_style(fp_p = normal_par, fp_t = normal_chunk),
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

  # Built-in heading styles. Names align with Word's "heading 1".."heading 3"
  # convention. outline_level wires the style into Word's navigation pane
  # and TOC field.
  x <- rtf_set_paragraph_style(
    x, style_id = "heading 1", style_name = "Heading 1",
    fp_p = fp_par(padding.top = 12, padding.bottom = 6),
    fp_t = fp_text_lite(bold = TRUE, font.size = 18, color = "#1F4E79"),
    outline_level = 1L
  )
  x <- rtf_set_paragraph_style(
    x, style_id = "heading 2", style_name = "Heading 2",
    fp_p = fp_par(padding.top = 8, padding.bottom = 4),
    fp_t = fp_text_lite(bold = TRUE, font.size = 14, color = "#2E75B6"),
    outline_level = 2L
  )
  x <- rtf_set_paragraph_style(
    x, style_id = "heading 3", style_name = "Heading 3",
    fp_p = fp_par(padding.top = 6, padding.bottom = 3),
    fp_t = fp_text_lite(bold = TRUE, font.size = 12, color = "#5B9BD5"),
    outline_level = 3L
  )
  x
}

#' @export
#' @title Add content into an RTF document
#' @description This function add 'officer' objects into an RTF document.
#' Values are added as new paragraphs. See section 'Methods (by class)'
#' that list supported objects.
#'
#' @section Section model in RTF:
#' A `block_section` added with `rtf_add()` applies to the content that
#' **follows** the call: it opens a new section whose layout (orientation,
#' columns, margins, headers / footers) is inherited by every paragraph,
#' table or graphic added afterwards, until the next `block_section` (or
#' the end of the document).
#'
#' Typical pattern: declare the section, then add the content.
#'
#' ```r
#' doc <- rtf_doc() |>
#'   rtf_add(block_section(prop_section(
#'     page_size = page_size(orient = "landscape")
#'   ))) |>
#'   rtf_add(fpar("This paragraph is in landscape orientation."))
#' ```
#'
#' The first section of the document is configured via the `def_sec`
#' argument of [rtf_doc()] (a [prop_section] object, not a
#' `block_section`). It applies to every element added before the first
#' `rtf_add(block_section(...))` call.
#'
#' The Word output uses the opposite model: `body_end_block_section()`
#' applies to the content that *precedes* the call. See
#' [body_end_block_section()].
#'
#' @param x rtf object, created by [rtf_doc()].
#' @param value object to add in the document. Supported objects
#' are vectors, graphics, block of formatted paragraphs. Use package
#' 'flextable' to add tables.
#' @param ... further arguments passed to or from other methods. When
#' adding a `ggplot` object or [plot_instr], these arguments will be used
#' by png function. See section 'Methods' to see what arguments can be used.
#' @example inst/examples/example_rtf_add.R
#' @example inst/examples/example_rtf_sections.R
rtf_add <- function(x, value, ...) {
  UseMethod("rtf_add", value)
}


#' @export
#' @describeIn rtf_add add a new section definition
rtf_add.block_section <- function(x, value, ...) {
  x$content[[length(x$content) + 1]] <- value
  x
}

# Walk the based_on chain of a style, returning the style ids in root-to-leaf
# order. `parents` is a named vector that maps each style_id to its parent
# style_id (or NA for roots), used as a constant-time dictionary. `visited`
# carries ids already traversed, so the walk terminates when a cycle exists
# (a parent that loops back to a descendant).
walk_based_on_chain <- function(style_id, parents, visited = character()) {
  if (is.na(style_id) || style_id %in% visited) {
    return(character())
  }
  c(
    walk_based_on_chain(parents[[style_id]], parents, c(visited, style_id)),
    style_id
  )
}

# Resolve a user-facing style id to (rtf_index, rtf_props). Concatenates the
# fragments of the based_on chain so the returned props string fully describes
# the style (the additive model the RTF spec requires for styled paragraphs).
resolve_rtf_style <- function(x, style) {
  if (is.null(style)) {
    return(NULL)
  }
  if (!is.character(style) || length(style) != 1L) {
    stop("`style` must be NULL or a single string matching a style_id.")
  }
  idx <- match(style, x$styles$style_id)
  if (is.na(idx)) {
    stop(
      "style ", shQuote(style),
      " is not registered. Use rtf_styles_info() to list available styles ",
      "or rtf_set_paragraph_style() to add one."
    )
  }
  # Walk the based_on chain from root to leaf. styles$rtf is stored
  # without leading \pard (see rtf_set_paragraph_style), so we can just
  # concatenate the chain in order: parent props come first, child
  # props win on conflicts under the RTF additive model.
  parents <- setNames(x$styles$based_on, x$styles$style_id)
  chain <- walk_based_on_chain(style, parents)
  rtf_by_id <- setNames(x$styles$rtf, x$styles$style_id)
  props <- paste0(rtf_by_id[chain], collapse = "")
  list(index = x$styles$rtf_index[idx], props = props)
}

apply_rtf_style <- function(value, style) {
  if (is.null(style)) {
    return(value)
  }
  attr(value, "rtf_style_index") <- style$index
  attr(value, "rtf_style_props") <- style$props
  value
}

#' @export
#' @describeIn rtf_add add characters as new paragraphs
#' @param style style identifier (`style_id`) of a paragraph style registered
#' on the document. Defaults to `NULL` (use the document's normal style). See
#' [rtf_set_paragraph_style()] and [rtf_styles_info()].
rtf_add.character <- function(x, value, style = NULL, ...) {
  styled <- resolve_rtf_style(x, style)
  chunk_prop <- if (is.null(styled)) x$normal_chunk else NULL
  par_prop <- if (is.null(styled)) x$normal_par else fp_par()
  values <- lapply(
    X = str_encode_to_rtf(value),
    function(z) {
      apply_rtf_style(
        fpar(ftext(z, prop = chunk_prop), fp_p = par_prop),
        styled
      )
    }
  )

  x$content <- append(x$content, values)
  x
}

#' @export
#' @describeIn rtf_add add a factor vector as new paragraphs
rtf_add.factor <- function(x, value, style = NULL, ...) {
  styled <- resolve_rtf_style(x, style)
  chunk_prop <- if (is.null(styled)) x$normal_chunk else NULL
  par_prop <- if (is.null(styled)) x$normal_par else fp_par()
  values <- lapply(
    X = str_encode_to_rtf(as.character(value)),
    function(z) {
      apply_rtf_style(
        fpar(ftext(z, prop = chunk_prop), fp_p = par_prop),
        styled
      )
    }
  )

  x$content <- append(x$content, values)
  x
}

#' @export
#' @describeIn rtf_add add a double vector as new paragraphs
#' @param formatter function used to format the numerical values
rtf_add.double <- function(x, value, formatter = formatC, style = NULL, ...) {
  styled <- resolve_rtf_style(x, style)
  chunk_prop <- if (is.null(styled)) x$normal_chunk else NULL
  par_prop <- if (is.null(styled)) x$normal_par else fp_par()
  values <- lapply(
    X = value,
    function(z) {
      apply_rtf_style(
        fpar(ftext(formatC(z), prop = chunk_prop), fp_p = par_prop),
        styled
      )
    }
  )

  x$content <- append(x$content, values)
  x
}

#' @export
#' @describeIn rtf_add add an [fpar()]
rtf_add.fpar <- function(x, value, style = NULL, ...) {
  styled <- resolve_rtf_style(x, style)
  x$content[[length(x$content) + 1]] <- apply_rtf_style(value, styled)
  x
}

#' @export
#' @describeIn rtf_add add an [block_list()]
rtf_add.block_list <- function(x, value, style = NULL, ...) {
  styled <- resolve_rtf_style(x, style)
  items <- lapply(as.list(value), apply_rtf_style, style = styled)
  x$content <- append(x$content, items)
  x
}

#' @export
#' @describeIn rtf_add add a [block_toc()] (table of contents).
#' Word populates the TOC at open time from paragraphs styled with the
#' built-in heading styles (which carry an outline level). LibreOffice
#' will not render the TOC automatically.
rtf_add.block_toc <- function(x, value, ...) {
  x$content[[length(x$content) + 1]] <- value
  x
}


#' @export
#' @describeIn rtf_add add a ggplot2
#' @param width height in inches
#' @param height height in inches
#' @param res resolution of the png image in ppi
#' @param scale Multiplicative scaling factor, same as in ggsave
#' @param ppr [fp_par()] to apply to paragraph.
rtf_add.gg <- function(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
) {
  if (!requireNamespace("ggplot2")) {
    stop("package ggplot2 is required to use this function")
  }

  file <- tempfile(fileext = ".png")
  agg_png(
    filename = file,
    width = width,
    height = height,
    units = "in",
    res = res,
    scaling = scale,
    background = "transparent",
    ...
  )
  tryCatch(
    {
      print(value)
    },
    finally = {
      dev.off()
    }
  )

  value <- fpar(
    external_img(
      src = file,
      width = width,
      height = height,
      guess_size = FALSE
    ),
    fp_p = ppr
  )

  x$content[[length(x$content) + 1]] <- value
  x
}
#' @export
#' @describeIn rtf_add add a [plot_instr()] object
rtf_add.plot_instr <- function(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
) {
  file <- tempfile(fileext = ".png")
  agg_png(
    filename = file,
    width = width,
    height = height,
    units = "in",
    res = res,
    scaling = scale,
    background = "transparent",
    ...
  )
  tryCatch(
    {
      eval(value$code)
    },
    finally = {
      dev.off()
    }
  )

  value <- fpar(
    external_img(
      src = file,
      width = width,
      height = height,
      guess_size = FALSE
    ),
    fp_p = ppr
  )

  x$content[[length(x$content) + 1]] <- value
  x
}

#' @title Write an 'RTF' File
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

  # Walk content elements, emit one RTF chunk per element and interleave
  # per-section header/footer groups so that each section carries its own
  # {\header}{\footer}. The default section's hfr applies to content added
  # before the first block_section; each subsequent block_section opens a
  # new section whose hfr is emitted just before the next \sect (or at
  # end of document for the last section). hfr groups must precede the
  # \sect that closes the section they describe (RTF spec).
  default_sect <- to_rtf(x$default_section)
  rtf_content <- character()
  current_hfr <- attr(default_sect, "hfr")
  for (elem in x$content) {
    if (inherits(elem, "block_section")) {
      props <- to_rtf(elem$property)
      rtf_content <- c(
        rtf_content,
        current_hfr,
        paste0("\\sect\\sectd", props)
      )
      current_hfr <- attr(props, "hfr")
    } else {
      rtf_content <- c(rtf_content, to_rtf(elem))
    }
  }
  trailing_hfr <- current_hfr

  rtf_ss <- rtf_stylesheet(x)
  txt_to_scan <- c(rtf_content, rtf_ss, trailing_hfr)
  m <- gregexec(pattern = "%font:[a-zA-Z ]+%", text = txt_to_scan)
  fonts <- unique(unlist(regmatches(txt_to_scan, m)))
  fonts <- gsub("(^%font:|%$)", "", fonts)
  family_table <- font_table(families = fonts)

  m <- gregexec(pattern = "\\%ftcolor\\:[a-zA-Z #0-9]+\\%", text = txt_to_scan)
  ftcolor <- unlist(regmatches(txt_to_scan, m))
  m <- gregexec(
    pattern = "\\%ftshading\\:[a-zA-Z #0-9]+\\%",
    text = txt_to_scan
  )
  ftshading <- unlist(regmatches(txt_to_scan, m))
  m <- gregexec(
    pattern = "\\%ftlinecolor\\:[a-zA-Z #0-9]+\\%",
    text = txt_to_scan
  )
  ftlinecolor <- unlist(regmatches(txt_to_scan, m))
  m <- gregexec(
    pattern = "\\%ftbgcolor\\:[a-zA-Z #0-9]+\\%",
    text = txt_to_scan
  )
  ftbgcolor <- unlist(regmatches(txt_to_scan, m))

  supported_highlight <- c(
    "black",
    "blue",
    "cyan",
    "green",
    "magenta",
    "red",
    "yellow",
    "darkblue",
    "darkcyan",
    "darkgreen",
    "darkmagenta",
    "darkred",
    "#8B8000",
    "darkgrey",
    "lightgray"
  )

  colors <- unique(c(
    ftcolor,
    ftshading,
    ftlinecolor,
    ftbgcolor,
    supported_highlight
  ))
  colors <- gsub(
    "(^%(ftcolor|ftshading|ftlinecolor|ftbgcolor):|%$)",
    "",
    colors
  )
  color_table <- color_table(colors = colors)

  header <- c(
    "{\\rtf1\\deff0",
    rtfize_font_table(family_table),
    rtfize_color_table(color_table)
  )

  all_strings <- c(
    header,
    rtf_ss,
    default_sect,
    rtf_content,
    trailing_hfr,
    "}"
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
      is_ext_ansi <- intv < 128
      is_unic2b <- !is_ext_ansi & intv < 32768
      is_unic4b <- !is_ext_ansi & !is_unic2b
      charv[is_unic2b] <- paste0("\\uc1\\u", intv[is_unic2b], "?")
      charv[is_unic4b] <- paste0("\\uc1\\u", intv[is_unic4b] - 65536, "?")
      paste0(charv, collapse = "")
    },
    char_list,
    int_list,
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
    color[1],
    color[2],
    color[3]
  )
}

rtf_fp_tab <- function(x) {
  tab_style <- "\\tqdec"
  if (x$style %in% "left") {
    tab_style <- ""
  } else if (x$style %in% "center") {
    tab_style <- "\\tqc"
  } else if (x$style %in% "right") {
    tab_style <- "\\tqr"
  }

  sprintf(
    "\\tx%.0f%s",
    inch_to_tweep(x$pos),
    tab_style
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
    "black",
    "blue",
    "cyan",
    "green",
    "magenta",
    "red",
    "yellow",
    "darkblue",
    "darkcyan",
    "darkgreen",
    "darkmagenta",
    "darkred",
    "#8B8000",
    "darkgrey",
    "lightgray"
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
      if (
        !used_shading %in%
          c(supported_highlight$match_ftshading, "%ftshading:transparent%")
      ) {
        stop(
          "chunk highlight color (",
          used_shading,
          ") in RTF can only have values: ",
          paste0(shQuote(supported_highlight$colors), collapse = ", ")
        )
      }
    }

    regmatches(x, m) <- sprintf("\\highlight%.0f", color_tbl$index[i])
  }
  x <- gsub("%ftshading:transparent%", "", x, fixed = TRUE)
  x
}

rtf_par_style <- function(fp_p = fp_par(), fp_t = NULL, include_pard = FALSE) {
  if (!is.null(fp_t)) {
    fp_t_rtf <- rpr_rtf(fp_t)
  } else {
    fp_t_rtf <- ""
  }

  paste0(ppr_rtf(fp_p, include_pard = include_pard), fp_t_rtf)
}

#' @export
#' @title Add or replace a paragraph style in an RTF document
#' @description Add or replace a paragraph style in the document stylesheet
#' so that subsequent paragraphs can reference it via the `style` argument of
#' [rtf_add()]. The function mirrors [docx_set_paragraph_style()] for Word.
#' @param x an rtf object created by [rtf_doc()].
#' @param style_id user-facing identifier used as the key when a paragraph
#' references the style (`rtf_add(..., style = style_id)`).
#' @param style_name display label of the style as it appears in Word's style
#' menu. Defaults to `style_id`.
#' @param base_on the `style_id` of the parent style. Properties not defined
#' in the new style are inherited from the parent. Defaults to `"Normal"`.
#' @param fp_p paragraph formatting properties, see [fp_par()].
#' @param fp_t text formatting properties, see [fp_text()]. If `NULL` the
#' paragraph inherits text properties from the parent style.
#' @param outline_level integer between 1 and 9, or `NULL`. Sets the outline
#' level (`\\outlinelevel`) so paragraphs using the style feed Word's
#' navigation pane and TOC field. Level 1 is the topmost (used by Heading 1).
#' @return the rtf object with the style added or updated.
#' @seealso [docx_set_paragraph_style()], [rtf_doc()], [rtf_add()]
#' @examples
#' doc <- rtf_doc()
#' doc <- rtf_set_paragraph_style(
#'   doc,
#'   style_id = "Callout",
#'   fp_p = fp_par(text.align = "center", padding = 6),
#'   fp_t = fp_text_lite(bold = TRUE, color = "#1F4E79")
#' )
#' doc <- rtf_add(doc, "Heads up", style = "Callout")
#' print(doc, target = tempfile(fileext = ".rtf"))
rtf_set_paragraph_style <- function(
  x,
  style_id,
  style_name = style_id,
  base_on = "Normal",
  fp_p = fp_par(),
  fp_t = NULL,
  outline_level = NULL
) {
  stopifnot(is.character(style_id), length(style_id) == 1L, nzchar(style_id))
  stopifnot(is.character(style_name), length(style_name) == 1L)
  if (!is.null(base_on)) {
    if (!base_on %in% x$styles$style_id) {
      stop(
        "base_on = ", shQuote(base_on),
        " does not match any existing style_id. Existing: ",
        paste0(shQuote(x$styles$style_id), collapse = ", ")
      )
    }
  }
  outline_str <- ""
  if (!is.null(outline_level)) {
    stopifnot(
      is.numeric(outline_level), length(outline_level) == 1L,
      outline_level >= 1L, outline_level <= 9L
    )
    # RTF / OOXML outline levels are 0-indexed; Heading 1 → \outlinelevel0.
    outline_str <- sprintf("\\outlinelevel%.0f", as.integer(outline_level) - 1L)
  }
  # Place \outlinelevelN inside the paragraph properties section, before
  # run-level properties, otherwise some readers stop parsing the
  # property list. include_pard = FALSE because \pard inside a style
  # definition (or duplicated in a styled paragraph after \pard\plain)
  # terminates the style's property list in Word's parser.
  ppr <- ppr_rtf(fp_p, include_pard = FALSE)
  rpr <- if (!is.null(fp_t)) rpr_rtf(fp_t) else ""
  rtf_fragment <- paste0(ppr, outline_str, rpr)

  idx <- which(x$styles$style_id == style_id)
  if (length(idx) < 1L) {
    rtf_index <- max(x$styles$rtf_index, 0L) + 1L
    new_row <- data.frame(
      style_id = style_id,
      style_name = style_name,
      type = "paragraph",
      rtf_index = rtf_index,
      based_on = if (is.null(base_on)) NA_character_ else base_on,
      rtf = rtf_fragment,
      stringsAsFactors = FALSE
    )
    x$styles <- rbind(x$styles, new_row)
  } else {
    x$styles[idx, "style_name"] <- style_name
    x$styles[idx, "based_on"] <- if (is.null(base_on)) NA_character_ else base_on
    x$styles[idx, "rtf"] <- rtf_fragment
  }
  x
}

#' @export
#' @title Read paragraph styles defined on an RTF document
#' @description Return the data.frame of styles currently registered on
#' the document. Useful for debugging or for checking which built-in
#' styles are available before referencing them via `style = ...`.
#' @param x an rtf object created by [rtf_doc()].
#' @return a data.frame with one row per style.
#' @seealso [rtf_set_paragraph_style()], [styles_info()]
#' @examples
#' rtf_styles_info(rtf_doc())
rtf_styles_info <- function(x) {
  stopifnot(inherits(x, "rtf"))
  x$styles
}

rtf_stylesheet <- function(x) {
  styles <- x$styles
  out <- character(nrow(styles))
  for (i in seq_len(nrow(styles))) {
    based_on_str <- ""
    if (!is.na(styles$based_on[i])) {
      parent_idx <- styles$rtf_index[match(styles$based_on[i], styles$style_id)]
      if (!is.na(parent_idx)) {
        based_on_str <- sprintf("\\sbasedon%.0f", parent_idx)
      }
    }
    # RTF spec canonical order in a paragraph style definition:
    #   \sN <paragraph props> <character props> <metadata> Name;
    # \sbasedon belongs to the metadata trailer; placing it right
    # after \sN makes Word attribute the property keywords that
    # follow to the parent style. The fragment in `styles$rtf` is
    # already emitted without \pard (rtf_set_paragraph_style uses
    # ppr_rtf(include_pard = FALSE)).
    out[i] <- sprintf(
      "\n{\\s%.0f%s%s %s;}",
      styles$rtf_index[i],
      styles$rtf[i],
      based_on_str,
      styles$style_name[i]
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
  gregexec <- function(
    pattern,
    text,
    ignore.case = FALSE,
    perl = FALSE,
    fixed = FALSE,
    useBytes = FALSE
  ) {
    if (is.factor(text) && nlevels(text) < length(text)) {
      out <- gregexec(
        pattern,
        c(levels(text), NA_character_),
        ignore.case,
        perl,
        fixed,
        useBytes
      )
      outna <- out[length(out)]
      out <- out[text]
      out[is.na(text)] <- outna
      return(out)
    }

    dat <- gregexpr(
      pattern = pattern,
      text = text,
      ignore.case = ignore.case,
      fixed = fixed,
      useBytes = useBytes,
      perl = perl
    )
    if (perl && !fixed) {
      ## Perl generates match data, so use that
      capt.attr <- c("capture.start", "capture.length", "capture.names")
      process <- function(x) {
        if (anyNA(x) || any(x < 0)) {
          y <- x
        } else {
          ## Interleave matches with captures
          y <- t(cbind(x, attr(x, "capture.start")))
          attributes(y)[names(attributes(x))] <- attributes(x)
          ml <- t(cbind(attr(x, "match.length"), attr(x, "capture.length")))
          nm <- attr(x, "capture.names")
          ## Remove empty names that `gregexpr` returns
          dimnames(ml) <- dimnames(y) <-
            if (any(nzchar(nm))) list(c("", nm), NULL)
          attr(y, "match.length") <- ml
          y
        }
        attributes(y)[capt.attr] <- NULL
        y
      }
      lapply(dat, process)
    } else {
      ## For TRE or fixed we must compute the match data ourselves
      m1 <- lapply(
        regmatches(text, dat),
        regexec,
        pattern = pattern,
        ignore.case = ignore.case,
        perl = perl,
        fixed = fixed,
        useBytes = useBytes
      )
      mlen <- lengths(m1)
      res <- vector("list", length(m1))
      im <- mlen > 0
      res[!im] <- dat[!im] # -1, NA
      res[im] <- Map(
        function(outer, inner) {
          tmp <- do.call(cbind, inner)
          attributes(tmp)[names(attributes(inner))] <- attributes(inner)
          attr(tmp, "match.length") <-
            do.call(cbind, lapply(inner, `attr`, "match.length"))
          # useBytes/index.type should be same for all so use outer vals
          attr(tmp, "useBytes") <- attr(outer, "useBytes")
          attr(tmp, "index.type") <- attr(outer, "index.type")
          tmp + rep(outer - 1L, each = nrow(tmp))
        },
        dat[im],
        m1[im]
      )
      res
    }
  }
}
