complete_styles_mapping <- function(
  styles_id_from,
  style_mapping,
  styles_info_tbl_to,
  styles_info_tbl_from,
  style_type = "paragraph",
  document_part_label
) {
  if (length(styles_id_from) < 1L) {
    return(style_mapping)
  }

  style_data_to <- styles_info_tbl_to[
    styles_info_tbl_to$style_type %in% style_type,
  ]
  style_data_from <- styles_info_tbl_from[
    styles_info_tbl_from$style_type %in% style_type,
  ]

  def_name_to <- head(style_data_to$style_name[style_data_to$is_default], n = 1)

  styles_name_from <- style_data_from$style_name[
    match(
      styles_id_from,
      style_data_from$style_id,
      nomatch = head(which(style_data_from$is_default), n = 1)
    )
  ]

  missing_styles <- setdiff(styles_name_from, unlist(style_mapping))

  if (length(missing_styles) > 0) {
    possible_match <- style_data_to$style_name[
      style_data_to$style_name %in% missing_styles
    ]
    for (amatch in possible_match) {
      if (is.null(style_mapping[[amatch]])) {
        style_mapping[[amatch]] <- character()
      }
      style_mapping[[amatch]] <- append(style_mapping[[amatch]], missing_styles)
    }
    missing_styles <- setdiff(missing_styles, possible_match)
  }

  if (length(missing_styles) > 0) {
    missing_styles_cli <- style_data_from[
      style_data_from$style_name %in% missing_styles,
      c("style_name", "style_id")
    ]
    missing_styles_cli <- paste0(
      "Style ",
      shQuote(missing_styles_cli$style_name),
      " with id {.code ",
      missing_styles_cli$style_id,
      "} could not be mapped"
    )

    names(missing_styles_cli) <- rep("*", length(missing_styles_cli))
    cli::cli_warn(
      c(
        "!" = "Style(s) mapping(s) for '{style_type}s' are missing in the {document_part_label} of the document:",
        missing_styles_cli,
        "i" = "Style name '{def_name_to}' is used instead."
      )
    )

    if (is.null(style_mapping[[def_name_to]])) {
      style_mapping[[def_name_to]] <- character()
    }
    style_mapping[[def_name_to]] <- append(
      style_mapping[[def_name_to]],
      missing_styles
    )
  }

  style_mapping
}

style_mapping_fortify <- function(
  style_mapping,
  styles_ids,
  styles_info_tbl_from,
  styles_info_tbl_to,
  style_type = "paragraph",
  document_part_label
) {
  style_mapping <- complete_styles_mapping(
    style_mapping = style_mapping,
    styles_id_from = styles_ids,
    styles_info_tbl_to = styles_info_tbl_to,
    styles_info_tbl_from = styles_info_tbl_from,
    style_type = style_type,
    document_part_label = document_part_label
  )

  style_mapping <- mapply(
    function(name_from, name_to) {
      data.frame(
        name_from = name_from,
        name_to = rep(name_to, length(name_from))
      )
    },
    name_from = style_mapping,
    name_to = names(style_mapping),
    SIMPLIFY = FALSE
  )
  style_mapping <- do.call(rbind, style_mapping)
  row.names(style_mapping) <- NULL
  if (is.null(style_mapping)) {
    style_mapping <- data.frame(
      name_from = character(0),
      name_to = character(0)
    )
  }

  sty_par_info_from <- styles_info_tbl_from[
    styles_info_tbl_from$style_type %in% style_type,
    c("style_id", "style_name")
  ]
  sty_par_info_to <- styles_info_tbl_to[
    styles_info_tbl_to$style_type %in% style_type,
    c("style_id", "style_name")
  ]

  style_mapping_par <- merge(
    x = style_mapping,
    y = sty_par_info_to,
    by.x = "name_to",
    by.y = "style_name",
    all.x = FALSE,
    all.y = FALSE
  )
  names(style_mapping_par)[names(style_mapping_par) %in% "style_id"] <- "id_to"
  style_mapping_par <- merge(
    x = style_mapping_par,
    y = sty_par_info_from,
    by.x = "name_from",
    by.y = "style_name",
    all.x = TRUE,
    all.y = FALSE
  )
  names(style_mapping_par)[
    names(style_mapping_par) %in% "style_id"
  ] <- "id_from"
  style_mapping_par <- style_mapping_par[c("id_from", "id_to")]

  style_mapping_par
}

numberings_append_xml <- function(file_numbering_from, file_numbering_to) {
  numbering_from <- read_xml(file_numbering_from)
  abstractnum_from <- xml_find_all(numbering_from, "w:abstractNum")
  num_from <- xml_find_all(numbering_from, "w:num")

  numbering_to <- read_xml(file_numbering_to)
  abstractnum_to <- xml_find_all(numbering_to, "w:abstractNum")
  num_to <- xml_find_all(numbering_to, "w:num")

  used_num_id_to <- xml_attr(num_to, "numId")
  used_num_id_to <- as.integer(used_num_id_to)
  used_num_id_to_next <- max(used_num_id_to) + 1L

  abst_num_id_to <- xml_attr(abstractnum_to, "abstractNumId")
  abst_num_id_to <- as.integer(abst_num_id_to)
  abst_num_id_to_next <- max(abst_num_id_to) + 1L

  used_num_id_from <- xml_attr(num_from, "numId")
  used_num_id_from <- as.integer(used_num_id_from)
  used_num_id_from_next <- max(used_num_id_from) + 1L

  abst_num_id_from <- xml_attr(abstractnum_from, "abstractNumId")
  abst_num_id_from <- as.integer(abst_num_id_from)
  abst_num_id_from_next <- max(abst_num_id_from) + 1L

  for (id_from in abst_num_id_from) {
    abstractNumFrom <- xml_find_first(
      numbering_from,
      sprintf(
        "w:abstractNum[@w:abstractNumId='%s']",
        as.character(id_from)
      )
    )
    xml_attr(abstractNumFrom, "w:abstractNumId") <- as.character(
      abst_num_id_to_next
    )

    usedAbsNumFrom <- xml_find_all(
      numbering_from,
      sprintf(
        "w:num/w:abstractNumId[@w:val='%s']",
        as.character(id_from)
      )
    )
    xml_attr(usedAbsNumFrom, "w:val") <- as.character(abst_num_id_to_next)
    abst_num_id_to_next <- abst_num_id_to_next + 1L
  }

  mapping_from <- data.frame(
    from = used_num_id_from,
    to = seq(used_num_id_to_next, along.with = used_num_id_from, by = 1L)
  )

  for (i in seq_len(nrow(mapping_from))) {
    id_from <- mapping_from$from[i]
    id_to <- mapping_from$to[i]
    usedNumFrom <- xml_find_all(
      numbering_from,
      sprintf("w:num[@w:numId='%s']", as.character(id_from))
    )
    xml_attr(usedNumFrom, "w:numId") <- as.character(id_to)
  }

  for (node in abstractnum_from) {
    xml_add_sibling(
      xml_find_first(numbering_to, "w:abstractNum"),
      node
    )
  }
  for (node in num_from) {
    xml_add_sibling(
      xml_find_first(numbering_to, "w:num"),
      node
    )
  }

  write_xml(
    numbering_to,
    file = file_numbering_to
  )

  mapping_from
}

# docx_part -----
docx_part <- R6Class(
  "docx_part",
  inherit = openxml_document,
  public = list(
    initialize = function(path, main_file, cursor, body_xpath) {
      super$initialize("word")
      private$package_dir <- path
      private$body_xpath <- body_xpath
      super$feed(file.path(private$package_dir, "word", main_file))
      private$cursor <- cursor
    },
    length = function() {
      xml_length(xml_find_first(self$get(), private$body_xpath))
    },
    patch_wml = function(
      package_dir,
      styles_info_tbl_from,
      styles_info_tbl_to,
      numbering_mapping,
      par_style_mapping = list(),
      run_style_mapping = list(),
      tbl_style_mapping = list(),
      additional_ns = character(0),
      prepend_chunks_on_styles = list()
    ) {
      doc_str <- self$encode_wml_str(additional_ns = additional_ns)
      document_rels <- self$rel_df()

      # images processing -----
      doc_from_img <- document_rels[basename(document_rels$type) %in% "image", ]
      for (i in seq_len(nrow(doc_from_img))) {
        fileext <- paste0(".", tools::file_ext(doc_from_img$target[i]))
        new_file <- tempfile(fileext = fileext)
        file.copy(
          from = file.path(
            package_dir,
            "word/media",
            basename(doc_from_img$target[i])
          ),
          to = new_file,
          overwrite = TRUE
        )
        pat <- "r:embed=\"%s\""
        pat <- sprintf(pat, doc_from_img$id[i])
        m <- gregexpr(pat, doc_str)
        regmatches(doc_str, m) <- sprintf(
          "r:embed=\"%s\"",
          new_file
        )
      }

      # external links processing -----
      doc_from_hl <- document_rels[
        basename(document_rels$type) %in% "hyperlink",
      ]
      for (i in seq_len(nrow(doc_from_hl))) {
        pat <- "r:id=\"%s\""
        pat <- sprintf(pat, doc_from_hl$id[i])
        m <- gregexpr(pat, doc_str)
        regmatches(doc_str, m) <- sprintf("r:id=\"%s\"", doc_from_hl$target[i])
      }

      # numberings processing -----
      for (i in seq_len(nrow(numbering_mapping))) {
        id_from <- numbering_mapping$from[i]
        id_to <- numbering_mapping$to[i]
        m <- gregexpr(
          sprintf("<w:numId w:val=\"%s\"/>", as.character(id_from)),
          doc_str,
          fixed = TRUE
        )
        regmatches(doc_str, m) <- sprintf(
          "<w:numId w:val=\"%s\"/>",
          as.character(id_to)
        )
      }

      sty_par_info_to <- styles_info_tbl_to[
        styles_info_tbl_to$style_type %in% "paragraph",
      ]
      sty_par_info_from <- styles_info_tbl_from[
        styles_info_tbl_from$style_type %in% "paragraph",
      ]

      # append chunks when specific styles are found -----
      for(style in names(prepend_chunks_on_styles)) {
        style_id <- sty_par_info_from$style_id[sty_par_info_from$style_name %in% style]
        # find all paragraphs with this style
        match_pstyle <- grep(sprintf("w:pStyle w:val=\"%s\"", style_id), doc_str)
        # find all </w:pPr> after each match
        match_end_ppr <- grep("</w:pPr>", doc_str)
        for(par_i in match_pstyle) {
          # find next </w:pPr>
          current_match_end_ppr <- match_end_ppr[match_end_ppr > par_i]
          current_match_end_ppr <- head(current_match_end_ppr, n = 1)
          # prepend chunk
          if (length(current_match_end_ppr) == 1) {
            doc_str[current_match_end_ppr] <- paste0(
              doc_str[current_match_end_ppr],
              to_wml(prepend_chunks_on_styles[[style]])
            )
          }
        }
      }

      # par styles processing -----
      m <- gregexpr("w:pStyle w:val=\"[[:alnum:]]+\"", doc_str)
      zz <- regmatches(doc_str, m)
      zz <- unlist(zz)
      p_styles <- gsub("w:pStyle w:val=\"([[:alnum:]]+)\"", "\\1", zz)
      p_styles <- unique(p_styles)
      p_styles <- setdiff(p_styles, sty_par_info_to$style_id)

      mapping_styles <- style_mapping_fortify(
        styles_ids = p_styles,
        style_mapping = par_style_mapping,
        styles_info_tbl_from = styles_info_tbl_from,
        styles_info_tbl_to = styles_info_tbl_to,
        style_type = "paragraph",
        document_part_label = self$document_part_label()
      )
      for (i in seq_len(nrow(mapping_styles))) {
        doc_str <- gsub(
          paste0("w:pStyle w:val=\"", mapping_styles$id_from[i], "\""),
          sprintf("w:pStyle w:val=\"%s\"", mapping_styles$id_to[i]),
          doc_str
        )
      }

      # runs/characters styles processing -----
      sty_chr_info_to <- styles_info_tbl_to[
        styles_info_tbl_to$style_type %in% "character",
      ]

      m <- gregexpr("w:rStyle w:val=\"[[:alnum:]]+\"", doc_str)
      zz <- regmatches(doc_str, m)
      zz <- unlist(zz)
      r_styles <- gsub("w:rStyle w:val=\"([[:alnum:]]+)\"", "\\1", zz)
      r_styles <- unique(r_styles)
      r_styles <- setdiff(r_styles, sty_chr_info_to$style_id)

      mapping_styles <- style_mapping_fortify(
        styles_ids = r_styles,
        style_mapping = run_style_mapping,
        styles_info_tbl_from = styles_info_tbl_from,
        styles_info_tbl_to = styles_info_tbl_to,
        style_type = "character",
        document_part_label = self$document_part_label()
      )
      for (i in seq_len(nrow(mapping_styles))) {
        doc_str <- gsub(
          paste0("w:rStyle w:val=\"", mapping_styles$id_from[i], "\""),
          sprintf("w:rStyle w:val=\"%s\"", mapping_styles$id_to[i]),
          doc_str
        )
      }

      # tables styles processing -----
      sty_tab_info_to <- styles_info_tbl_to[
        styles_info_tbl_to$style_type %in% "table",
      ]

      m <- gregexpr("w:tblStyle w:val=\"[[:alnum:]]+\"", doc_str)
      zz <- regmatches(doc_str, m)
      zz <- unlist(zz)
      tbl_styles <- gsub("w:tblStyle w:val=\"([[:alnum:]]+)\"", "\\1", zz)
      tbl_styles <- unique(tbl_styles)
      tbl_styles <- setdiff(tbl_styles, sty_tab_info_to$style_id)
      mapping_styles <- style_mapping_fortify(
        styles_ids = tbl_styles,
        style_mapping = tbl_style_mapping,
        styles_info_tbl_from = styles_info_tbl_from,
        styles_info_tbl_to = styles_info_tbl_to,
        style_type = "table",
        document_part_label = self$document_part_label()
      )
      for (i in seq_len(nrow(mapping_styles))) {
        doc_str <- gsub(
          paste0("w:tblStyle w:val=\"", mapping_styles$id_from[i], "\""),
          sprintf("w:tblStyle w:val=\"%s\"", mapping_styles$id_to[i]),
          doc_str
        )
      }
      doc_str
    }
  ),
  private = list(
    package_dir = NULL,
    cursor = NULL,
    body_xpath = NULL
  )
)

# body_part -----
body_part <- R6Class(
  "body_part",
  inherit = docx_part,
  public = list(
    document_part_label = function() {
      "body"
    },
    encode_wml_str = function(additional_ns) {
      body <- self$get()
      body <- xml_find_first(body, "w:body")

      # sections are removed
      xml_remove(xml_find_all(body, "//w:sectPr"))

      # comments are removed
      xml_remove(xml_find_all(body, "//w:commentRangeStart"))
      xml_remove(xml_find_all(body, "//w:commentRangeEnd"))
      xml_remove(xml_find_all(body, "//w:commentReference"))

      chr_body <- xml_document_to_chrs(body)
      chr_body
    }
  )
)


# footnotes_part -----
footnotes_part <- R6Class(
  "footnotes_part",
  inherit = docx_part,
  public = list(
    document_part_label = function() {
      "footnotes"
    },
    encode_wml_str = function(additional_ns) {
      footnotes <- self$get()
      footnotes <- xml_find_all(footnotes, "w:footnote[not(@w:type)]")

      chr_footnotes <- vapply(
        footnotes,
        function(x) {
          str <- sapply(xml_children(x), as.character)
          paste0(str, collapse = "")
        },
        FUN.VALUE = ""
      )
      add_ns_str <- sprintf(
        " xmlns:%s=\"%s\"",
        names(additional_ns),
        additional_ns
      )
      add_ns_str <- paste0(add_ns_str, collapse = "")

      chr_footnotes <- paste0(
        "<w:footnoteReference w:officer=\"true\" w:id=\"%s\">",
        "<w:footnote ",
        add_ns_str,
        " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:id=\"%s\">",
        chr_footnotes,
        "</w:footnote>",
        "</w:footnoteReference>"
      )
      names(chr_footnotes) <- xml_attr(footnotes, "id")

      chr_footnotes
    }
  )
)
