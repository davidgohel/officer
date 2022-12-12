get_val <- function(x, attr) {
  xml_attr(xml_child(x, attr), "val")
}

read_rpr = function(x) {
  node_rpr <- xml_child(x, "w:rPr")

  if (inherits(node_rpr, "xml_missing")) {
    return(structure(list(.dummy = NA_character_), names = ".dummy", row.names = c(NA, -1L
    ), class = "data.frame"))
  }

  lang <- xml_child(node_rpr, "w:lang")
  fonts <- xml_child(node_rpr, "w:rFonts")
  shd <- xml_child(node_rpr, "w:shd")

  data.frame(
    stringsAsFactors = FALSE,
    font.size = get_val(node_rpr, "w:sz"),
    bold = get_val(node_rpr, "w:b"),
    italic = get_val(node_rpr, "w:i"),
    underlined = get_val(node_rpr, "w:u"),
    color = get_val(node_rpr, "w:color"),
    font.family = xml_attr(fonts, "ascii"),

    vertical.align = get_val(node_rpr, "w:vertAlign"),
    shading.color = xml_attr(shd, "fill"),

    hansi.family = xml_attr(fonts, "hAnsi"),
    eastasia.family = xml_attr(fonts, "eastAsia"),
    cs.family = xml_attr(fonts, "cs"),
    bold.cs = get_val(node_rpr, "w:bCs"),
    font.size.cs = get_val(node_rpr, "w:szCs"),

    lang.val = xml_attr(lang, "val"),
    lang.eastasia = xml_attr(lang, "eastAsia"),
    lang.bidi = xml_attr(lang, "bidi")
  )
}
read_ppr = function(x) {
  node_ppr <- xml_child(x, "w:pPr")

  if (inherits(node_ppr, "xml_missing")) {
    return(structure(list(.dummy = NA_character_), names = ".dummy", row.names = c(NA, -1L
    ), class = "data.frame"))
  }
  spacing <- xml_child(node_ppr, "w:spacing")
  indent <- xml_child(node_ppr, "w:ind")
  shd <- xml_child(node_ppr, "w:shd")

  bdr <- xml_child(node_ppr, "w:pBdr")
  border.bottom <- xml_child(bdr, "w:bottom")
  border.top <- xml_child(bdr, "w:top")
  border.left <- xml_child(bdr, "w:left")
  border.right <- xml_child(bdr, "w:right")

  spacing <- xml_child(node_ppr, "w:spacing")
  line_spacing <- as.integer(xml_attr(spacing, "line")) / 240

  data.frame(
    stringsAsFactors = FALSE,
    align = get_val(node_ppr, "w:jc"),
    keep_next = !inherits(xml_child(node_ppr, "w:keepNext"), "xml_missing"),
    line_spacing = line_spacing,
    padding.bottom = xml_attr(spacing, "after"),
    padding.top = xml_attr(spacing, "before"),
    padding.left = xml_attr(indent, "left"),
    padding.right= xml_attr(indent, "right"),
    shading.color.par = xml_attr(shd, "fill"),
    border.bottom.width = as.integer(xml_attr(border.bottom, "sz")) / 8,
    border.bottom.color = xml_attr(border.bottom, "color"),
    border.bottom.style = xml_attr(border.bottom, "val"),
    border.top.width = as.integer(xml_attr(border.top, "sz")) / 8,
    border.top.color = xml_attr(border.top, "color"),
    border.top.style = xml_attr(border.top, "val"),
    border.left.width = as.integer(xml_attr(border.left, "sz")) / 8,
    border.left.color = xml_attr(border.left, "color"),
    border.left.style = xml_attr(border.left, "val"),
    border.right.width = as.integer(xml_attr(border.right, "sz")) / 8,
    border.right.color = xml_attr(border.right, "color"),
    border.right.style = xml_attr(border.right, "val")
  )
}

read_docx_styles <- function(package_dir){
  styles_file <- file.path(package_dir, "word/styles.xml")
  doc <- read_xml(styles_file)

  all_styles <- xml_find_all(doc, "/w:styles/w:style")

  ppr <- lapply(
    xml_find_all(doc, "/w:styles/w:style"),
    read_ppr)
  ppr <- rbind.match.columns(ppr)
  ppr$.dummy <- NULL

  rpr <- lapply(
    xml_find_all(doc, "/w:styles/w:style"),
    read_rpr)
  rpr <- rbind.match.columns(rpr)
  rpr$.dummy <- NULL

  types <- xml_attr(all_styles, "type")
  main <- data.frame(stringsAsFactors = FALSE,
             style_type = types,
             style_id = xml_attr(all_styles, "styleId"),
             style_name = xml_attr(xml_child(all_styles, "w:name"), "val"),
             base_on = xml_attr(xml_child(all_styles, "w:basedOn"), "val"),
             is_custom = xml_attr(all_styles, "customStyle") %in% "1",
             is_default = xml_attr(all_styles, "default") %in% "1"
  )
  out <- cbind(main, ppr, rpr)
  out
}

get_style_id <- function(data, style, type ){
  ref <- data[data$style_type==type, ]

  if(!style %in% ref$style_name){
    t_ <- shQuote(ref$style_name, type = "sh")
    t_ <- paste(t_, collapse = ", ")
    t_ <- paste0("c(", t_, ")")
    stop("could not match any style named ", shQuote(style, type = "sh"), " in ", t_, call. = FALSE)
  }
  ref$style_id[ref$style_name == style]
}

