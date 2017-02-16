#' @importFrom xml2 xml_find_all xml_attr read_xml
#' @import magrittr
#' @importFrom tibble tibble
#' @export
#' @title read word styles
#' @description read word styles and returns a tabular description
#' @param path path to Word document
#' @examples
#' template_dir <- tempfile()
#' docx_file <- system.file(package = "officer", "template/template.docx")
#' unpack_folder( file = docx_file, folder = template_dir )
#' read_styles( path = template_dir )
read_styles <- function( path ){
  styles_file <- file.path(path, "word/styles.xml")
  doc <- read_xml(styles_file)

  all_styles <- xml_find_all(doc, "/w:styles/w:style")
  all_desc <- tibble(
    style_type = all_styles %>% xml_attr("type"),
    style_id = all_styles %>% xml_attr("styleId"),
    style_name = all_styles %>% xml_find_all("w:name") %>% xml_attr("val"),
    is_custom = all_styles %>% xml_attr("customStyle") %in% "1",
    is_default = all_styles %>% xml_attr("default") %in% "1"
  )

  all_desc
}



#' @importFrom xml2 xml_ns read_xml xml_find_all xml_name xml_text
#' @title core properties
#' @description read core properties of a Word file and returns a tabular description
#' @param path path to Word document
#' @examples
#' template_dir <- tempfile()
#' docx_file <- system.file(package = "officer", "template/template.docx")
#' unpack_folder( file = docx_file, folder = template_dir )
#' read_core( path = template_dir )
#' @export
read_core <- function( path ){
  core_file <- file.path(path, "docProps/core.xml")
  doc <- read_xml(core_file)

  all_ <- xml_find_all(doc, "/cp:coreProperties/*")
  ns_ <- xml_ns(doc)
  all_desc <- tibble(
    tag_ns = all_ %>% xml_name(ns = ns_),
    tag = all_ %>% xml_name(),
    value = all_ %>% xml_text()
  )

  all_desc
}

