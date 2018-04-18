read_docx_styles <- function( package_dir ){
  styles_file <- file.path(package_dir, "word/styles.xml")
  doc <- read_xml(styles_file)

  all_styles <- xml_find_all(doc, "/w:styles/w:style")

  all_desc <- data.frame(stringsAsFactors = FALSE,
                         style_type = xml_attr(all_styles, "type"),
                         style_id = xml_attr(all_styles, "styleId"),
                         style_name = xml_attr(xml_child(all_styles, "w:name"), "val"),
                         is_custom = xml_attr(all_styles, "customStyle") %in% "1",
                         is_default = xml_attr(all_styles, "default") %in% "1"
  )

  all_desc
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

