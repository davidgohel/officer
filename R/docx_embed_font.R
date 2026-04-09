#' @export
#' @title Embed Fonts in a Word Document
#' @description Copy TrueType or OpenType font files from the local
#' system into a Word document. The font data is stored inside the
#' `.docx` archive so that the document renders correctly when opened
#' on a system where the font is not installed.
#'
#' When only `font_family` is provided (without `regular`), the function
#' looks up font file paths on the local system via [gdtools::sys_fonts()]
#' and copies them into the document. A message shows the equivalent
#' explicit call for reproducibility.
#'
#' @section Font licensing:
#' Embedding a font in a document redistributes it. You must ensure that
#' the font license permits embedding. Fonts under the SIL Open Font
#' License (e.g. Liberation, Google Fonts) generally allow it. Many
#' commercial fonts restrict or prohibit embedding. Check the license
#' of the font before using this function.
#' @param x an rdocx object
#' @param font_family font family name as it will appear in the document.
#' When `regular` is not provided, this name is used to look up font
#' files via [gdtools::sys_fonts()].
#' @param regular path to the regular .ttf/.otf font file. If `NULL`,
#' font files are detected automatically via gdtools.
#' @param bold path to the bold .ttf/.otf font file (optional)
#' @param italic path to the italic .ttf/.otf font file (optional)
#' @param bold_italic path to the bold-italic .ttf/.otf font file (optional)
#' @return the rdocx object with embedded fonts
#' @seealso [gdtools::sys_fonts()], [gdtools::register_gfont()],
#' [docx_set_settings()]
#' @examplesIf require("gdtools", quietly = TRUE, character.only = TRUE)
#' # automatic detection (requires gdtools)
#' gdtools::register_liberationsans()
#' doc <- read_docx()
#' doc <- docx_embed_font(doc, font_family = "Liberation Sans")
#'
#' # # explicit paths (no gdtools needed)
#' # doc <- docx_embed_font(
#' #   doc,
#' #   font_family = "My Font",
#' #   regular = "path/to/font-regular.ttf",
#' #   bold = "path/to/font-bold.ttf"
#' # )
docx_embed_font <- function(
  x,
  font_family,
  regular = NULL,
  bold = NULL,
  italic = NULL,
  bold_italic = NULL
) {
  if (!inherits(x, "rdocx")) {
    cli::cli_abort("{.arg x} must be an rdocx object.")
  }

  if (is.null(regular)) {
    detected <- detect_font_files(font_family)
    regular <- detected$regular
    bold <- detected$bold %||% bold
    italic <- detected$italic %||% italic
    bold_italic <- detected$bold_italic %||% bold_italic
  }

  if (!file.exists(regular)) {
    cli::cli_abort("Font file {.file {regular}} does not exist.")
  }

  styles <- list(
    embedRegular = regular,
    embedBold = bold,
    embedItalic = italic,
    embedBoldItalic = bold_italic
  )
  styles <- Filter(Negate(is.null), styles)

  for (path in styles) {
    if (!file.exists(path)) {
      cli::cli_abort("Font file {.file {path}} does not exist.")
    }
  }

  package_dir <- x$package_dir
  fonts_dir <- file.path(package_dir, "word", "fonts")
  if (!dir.exists(fonts_dir)) {
    dir.create(fonts_dir, recursive = TRUE)
  }

  # determine next font file number
  existing <- list.files(fonts_dir, pattern = "^font\\d+\\.odttf$")
  if (length(existing) > 0) {
    nums <- as.integer(gsub("\\D", "", existing))
    next_num <- max(nums) + 1L
  } else {
    next_num <- 1L
  }

  # update fontTable.xml.rels
  rels_file <- file.path(package_dir, "word", "_rels", "fontTable.xml.rels")
  if (file.exists(rels_file)) {
    rels_doc <- read_xml(rels_file)
    existing_rels <- xml_find_all(
      rels_doc,
      "d1:Relationship",
      ns = xml_ns(rels_doc)
    )
    existing_ids <- xml_attr(existing_rels, "Id")
    if (length(existing_ids) > 0) {
      next_rid <- max(as.integer(gsub("\\D", "", existing_ids))) + 1L
    } else {
      next_rid <- 1L
    }
  } else {
    rels_dir <- file.path(package_dir, "word", "_rels")
    if (!dir.exists(rels_dir)) {
      dir.create(rels_dir, recursive = TRUE)
    }
    rels_doc <- read_xml(paste0(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    ))
    next_rid <- 1L
  }

  # update fontTable.xml
  font_table_file <- file.path(package_dir, "word", "fontTable.xml")
  font_table_doc <- read_xml(font_table_file)
  ns_w <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  ns_r <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  font_node <- xml_find_first(
    font_table_doc,
    sprintf("w:font[@w:name='%s']", font_family),
    ns = c(w = ns_w)
  )
  if (inherits(font_node, "xml_missing")) {
    font_xml <- sprintf(
      '<w:font xmlns:w="%s" xmlns:r="%s" w:name="%s"/>',
      ns_w,
      ns_r,
      font_family
    )
    xml_add_child(font_table_doc, read_xml(font_xml))
    font_node <- xml_find_first(
      font_table_doc,
      sprintf("w:font[@w:name='%s']", font_family),
      ns = c(w = ns_w)
    )
  }

  # process each style
  rel_type <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
  ns_pkg <- "http://schemas.openxmlformats.org/package/2006/relationships"

  for (embed_tag in names(styles)) {
    font_path <- styles[[embed_tag]]
    guid <- format_guid(uuid_generate())
    key <- guid_to_xor_key(guid)

    # obfuscate
    font_bytes <- readBin(font_path, "raw", file.info(font_path)$size)
    font_bytes[1:16] <- xor(font_bytes[1:16], key)
    font_bytes[17:32] <- xor(font_bytes[17:32], key)

    # write odttf
    odttf_name <- sprintf("font%d.odttf", next_num)
    writeBin(font_bytes, file.path(fonts_dir, odttf_name))
    next_num <- next_num + 1L

    # add relationship
    rid <- sprintf("rId%d", next_rid)
    rel_xml <- sprintf(
      '<Relationship xmlns="%s" Id="%s" Type="%s" Target="fonts/%s"/>',
      ns_pkg,
      rid,
      rel_type,
      odttf_name
    )
    xml_add_child(rels_doc, read_xml(rel_xml))
    next_rid <- next_rid + 1L

    # remove existing embed node of same type if present
    old_embed <- xml_child(font_node, paste0("w:", embed_tag))
    if (!inherits(old_embed, "xml_missing")) {
      xml_remove(old_embed)
    }

    # add embed node
    embed_xml <- sprintf(
      '<w:%s xmlns:w="%s" xmlns:r="%s" r:id="%s" w:fontKey="%s"/>',
      embed_tag,
      ns_w,
      ns_r,
      rid,
      guid
    )
    xml_add_child(font_node, read_xml(embed_xml))
  }

  write_xml(rels_doc, rels_file)
  write_xml(font_table_doc, font_table_file)

  # content type for odttf
  x$content_type$add_ext(
    "odttf",
    "application/vnd.openxmlformats-officedocument.obfuscatedFont"
  )

  # ensure embedTrueTypeFonts in settings
  x$settings$embed_true_type_fonts <- TRUE

  x
}

# Detect font files from gdtools::sys_fonts() for a given family
detect_font_files <- function(font_family) {
  if (!requireNamespace("gdtools", quietly = TRUE)) {
    cli::cli_abort(c(
      "Package {.pkg gdtools} is required to auto-detect font files.",
      "i" = "Install it with {.code install.packages(\"gdtools\")}",
      "i" = "Or provide explicit file paths via {.arg regular}, {.arg bold}, etc."
    ))
  }

  sysfonts <- gdtools::sys_fonts()
  matches <- sysfonts[sysfonts$family == font_family, ]

  if (nrow(matches) == 0L) {
    cli::cli_abort(c(
      "Font family {.val {font_family}} not found in system fonts.",
      "i" = "Register it first with {.code gdtools::register_gfont()} or {.code gdtools::register_liberationsans()}.",
      "i" = "Check available families with {.code unique(gdtools::sys_fonts()$family)}."
    ))
  }

  style_map <- list(
    regular = "Regular",
    bold = "Bold",
    italic = "Italic",
    bold_italic = "Bold Italic"
  )

  result <- list()
  for (nm in names(style_map)) {
    row <- matches[matches$style == style_map[[nm]], ]
    if (nrow(row) > 0L) {
      result[[nm]] <- row$path[1L]
    }
  }

  if (is.null(result$regular)) {
    cli::cli_abort(c(
      "No {.val Regular} style found for font family {.val {font_family}}.",
      "i" = "Provide the path explicitly via {.arg regular}."
    ))
  }

  # print equivalent explicit call
  arg_names <- c(
    regular = "regular",
    bold = "bold",
    italic = "italic",
    bold_italic = "bold_italic"
  )
  present <- Filter(Negate(is.null), result)
  code_args <- vapply(
    names(present),
    function(nm) {
      sprintf("  %s = \"%s\"", arg_names[[nm]], present[[nm]])
    },
    character(1L)
  )
  code <- paste0(
    "docx_embed_font(\n  x,\n",
    sprintf("  font_family = \"%s\",\n", font_family),
    paste(code_args, collapse = ",\n"),
    "\n)"
  )
  cli::cli_inform(
    "Embedding {.val {font_family}} with detected font files:"
  )
  message(code)

  result
}

# Format a UUID string as a Word fontKey GUID
format_guid <- function(uuid_str) {
  toupper(paste0("{", uuid_str, "}"))
}

# Convert a GUID string like {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
# to the 16-byte XOR key used for OOXML font obfuscation
guid_to_xor_key <- function(guid_str) {
  hex <- gsub("[{}-]", "", guid_str)
  bytes <- as.raw(strtoi(substring(hex, seq(1, 31, 2), seq(2, 32, 2)), 16L))
  # Windows GUID memory layout: groups 1-3 little-endian, groups 4-5 big-endian
  guid_mem <- c(rev(bytes[1:4]), rev(bytes[5:6]), rev(bytes[7:8]), bytes[9:16])
  rev(as.raw(guid_mem))
}
