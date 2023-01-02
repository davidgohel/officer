get_all_attrs <- function(nodeset) {
    # extract all children per node
    children <- lapply(nodeset, xml_children)
    # extract all node names per child
    name <- lapply(children, xml_name)
    # extract all attributes per child
    attr <- lapply(children, xml_attrs)
    # set attribute names
    for (i in seq_along(attr)) {
      names(attr[[i]]) <- name[[i]]
    }
    attr
}

get_attr_val <- function(lst, node, attr = NULL) {
  if (is.null(attr)) attr <- "val"
  if (!length(lst)) return(NA_character_)

  vapply(lst, FUN.VALUE = "", function(node_attr) {
    a <- node_attr[[node]][attr]
    if (is.null(a)) NA_character_ else a
  })
}

# a faster way for constructing a data.frame
as_data_frame <- function(x, ...) {
  stopifnot(length(unique(lengths(x))) == 1L)
  structure(x, ..., row.names = seq_along(x[[1L]]), class = "data.frame")
}

read_rpr <- function(x) {
  node_rpr <- xml_child(x, "w:rPr")

  # check if empty
  not_empty <- lengths(node_rpr) > 0L

  # all elements to extract
  elem <- c(
    "font.size", "bold", "italic", "underlined", "color", "font.family",
    "vertical.align", "shading.color",
    "hansi.family", "eastasia.family", "cs.family", "bold.cs", "font.size.cs",
    "lang.val", "lang.eastasia", "lang.bidi"
  )

  # init results
  rpr <- replicate(length(elem), list(rep(NA_character_, length(node_rpr))))
  names(rpr) <- elem

  # directly return if all are empty
  if (!any(not_empty)) return(as_data_frame(rpr))

  # further reduce to only work on non-empty nodes and extract all attributes of
  # every child
  rpr_attr <- get_all_attrs(node_rpr[not_empty])

  # replace with actual attribute values
  rpr[["font.size"]][not_empty]       <- get_attr_val(rpr_attr, "sz")
  rpr[["bold"]][not_empty]            <- get_attr_val(rpr_attr, "b")
  rpr[["italic"]][not_empty]          <- get_attr_val(rpr_attr, "i")
  rpr[["underlined"]][not_empty]      <- get_attr_val(rpr_attr, "u")
  rpr[["color"]][not_empty]           <- get_attr_val(rpr_attr, "color")
  rpr[["font.family"]][not_empty]     <- get_attr_val(rpr_attr, "rFonts", "ascii")

  rpr[["vertical.align"]][not_empty]  <- get_attr_val(rpr_attr, "vertAlign")
  rpr[["shading.color"]][not_empty]   <- get_attr_val(rpr_attr, "shd", "fill")

  rpr[["hansi.family"]][not_empty]    <- get_attr_val(rpr_attr, "rFonts", "hAnsi")
  rpr[["eastasia.family"]][not_empty] <- get_attr_val(rpr_attr, "rFonts", "eastAsia")
  rpr[["cs.family"]][not_empty]       <- get_attr_val(rpr_attr, "rFonts", "cs")
  rpr[["bold.cs"]][not_empty]         <- get_attr_val(rpr_attr, "bCs")
  rpr[["font.size.cs"]][not_empty]    <- get_attr_val(rpr_attr, "szCs")

  rpr[["lang.val"]][not_empty]        <- get_attr_val(rpr_attr, "lang")
  rpr[["lang.eastasia"]][not_empty]   <- get_attr_val(rpr_attr, "lang", "eastAsia")
  rpr[["lang.bidi"]][not_empty]       <- get_attr_val(rpr_attr, "lang", "bidi")

  as_data_frame(rpr)
}

read_ppr <- function(x) {
  node_ppr <- xml_child(x, "w:pPr")

  # check if empty
  not_empty <- lengths(node_ppr) > 0L

  # all elements to extract
  elem <- c(
    "align", "keep_next", "line_spacing",
    "padding.bottom", "padding.top", "padding.left", "padding.right", "shading.color.par",
    "border.bottom.width", "border.bottom.color", "border.bottom.style",
    "border.top.width", "border.top.color", "border.top.style",
    "border.left.width", "border.left.color", "border.left.style",
    "border.right.width", "border.right.color", "border.right.style"
  )

  # init results
  ppr <- replicate(length(elem), list(rep(NA_character_, length(node_ppr))))
  names(ppr) <- elem

  # init values for non-character columns
  ppr[paste0("border.", c("bottom", "top", "left", "right"), ".width")] <- list(rep(NA_integer_, length(node_ppr)))
  ppr["line_spacing"] <- list(rep(NA_real_, length(node_ppr)))
  ppr["keep_next"] <- list(rep(FALSE, length(node_ppr)))

  # directly return if all are empty
  if (!any(not_empty)) return(as_data_frame(ppr))

  # further reduce to only work on non-empty nodes and extract all attributes of
  # every child
  ppr_attr <- get_all_attrs(node_ppr[not_empty])

  # NOTE: 'w:pBdr' has to be treated specially since it is nested inside 'w:pPr'
  bdr <- xml_child(node_ppr[not_empty], "w:pBdr")
  not_empty_bdr <- lengths(bdr) > 0L
  bdr_attr <- get_all_attrs(bdr[not_empty_bdr])

  # replace with actual attribute values
  ppr[["align"]][not_empty] <- get_attr_val(ppr_attr, "jc")
  ppr[["keep_next"]][not_empty] <- vapply(ppr_attr, function(attrs) "keepNext" %in% names(attrs), logical(1L))
  ppr[["line_spacing"]][not_empty] <- as.integer(get_attr_val(ppr_attr, "spacing", "line")) / 240L

  ppr[["padding.bottom"]][not_empty] <- get_attr_val(ppr_attr, "spacing", "after")
  ppr[["padding.top"]][not_empty] <- get_attr_val(ppr_attr, "spacing", "before")
  ppr[["padding.left"]][not_empty] <- get_attr_val(ppr_attr, "ind", "left")
  ppr[["padding.right"]][not_empty] <- get_attr_val(ppr_attr, "ind", "right")
  ppr[["shading.color.par"]][not_empty] <- get_attr_val(ppr_attr, "shd", "fill")

  ppr[["border.bottom.width"]][not_empty][not_empty_bdr] <- as.integer(get_attr_val(bdr_attr, "bottom", "sz")) / 8L
  ppr[["border.bottom.color"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "bottom", "color")
  ppr[["border.bottom.style"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "bottom", "val")

  ppr[["border.top.width"]][not_empty][not_empty_bdr] <- as.integer(get_attr_val(bdr_attr, "top", "sz")) / 8L
  ppr[["border.top.color"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "top", "color")
  ppr[["border.top.style"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "top", "val")

  ppr[["border.left.width"]][not_empty][not_empty_bdr] <- as.integer(get_attr_val(bdr_attr, "left", "sz")) / 8L
  ppr[["border.left.color"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "left", "color")
  ppr[["border.left.style"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "left", "val")

  ppr[["border.right.width"]][not_empty][not_empty_bdr] <- as.integer(get_attr_val(bdr_attr, "right", "sz")) / 8L
  ppr[["border.right.color"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "right", "color")
  ppr[["border.right.style"]][not_empty][not_empty_bdr] <- get_attr_val(bdr_attr, "right", "val")

  as_data_frame(ppr)
}

read_docx_styles <- function(package_dir) {
  styles_file <- file.path(package_dir, "word/styles.xml")
  doc <- read_xml(styles_file)

  all_styles <- xml_find_all(doc, "/w:styles/w:style")

  ppr <- read_ppr(all_styles)
  rpr <- read_rpr(all_styles)

  main <- data.frame(stringsAsFactors = FALSE,
    style_type = xml_attr(all_styles, "type"),
    style_id = xml_attr(all_styles, "styleId"),
    style_name = xml_attr(xml_child(all_styles, "w:name"), "val"),
    base_on = xml_attr(xml_child(all_styles, "w:basedOn"), "val"),
    is_custom = xml_attr(all_styles, "customStyle") %in% "1",
    is_default = xml_attr(all_styles, "default") %in% "1"
  )
  out <- cbind(main, ppr, rpr)
  out
}

get_style_id <- function(data, style, type) {
  ref <- data[data$style_type == type, ]

  if (!style %in% ref$style_name) {
    t_ <- shQuote(ref$style_name, type = "sh")
    t_ <- paste(t_, collapse = ", ")
    t_ <- paste0("c(", t_, ")")
    stop("could not match any style named ", shQuote(style, type = "sh"), " in ", t_, call. = FALSE)
  }
  ref$style_id[ref$style_name == style]
}
