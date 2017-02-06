#' @export
#' @title add a title
#' @description add a title into a pptx object
#' @param x a pptx device
#' @param value a character
#' @param pos where to add the new element relative to the cursor,
#' one of "after", "before", "on".
#' @importFrom xml2 read_xml xml_find_first write_xml xml_add_sibling as_xml_document
set_title <- function( x, value ){

  slide <- x$slide$get_slide(x$cursor)

  layoutdata <- x$slideLayouts$get_data()
  layoutdata <- layoutdata[ basename( layoutdata$filename) %in% slide$layout_name(), ]

  layout_doc <- read_xml(file.path(x$package_dir, "ppt/slideLayouts", layoutdata$filename ))
  template_node <- xml_find_first(layout_doc, "//p:sp[p:nvSpPr/p:nvPr/p:ph[contains(@type,'title')]]")

  xml_remove( xml_find_all(template_node, "//p:txBody/a:p") )
  extra_ns <- "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""
  xml_elt <- sprintf( "<a:p %s><a:r><a:rPr/><a:t>%s</a:t></a:r></a:p>", extra_ns, value )
  xml_add_child( xml_find_first(template_node, "//p:txBody"), as_xml_document(xml_elt) )

  xml_list <- xml_find_first(slide$get(), "//p:sp[p:nvSpPr/p:nvPr/p:ph[contains(@type,'title')]]")

  if( !inherits(xml_list, "xml_missing")){
    xml_replace(xml_list, template_node)
  } else{
    xml_add_child(xml_find_first(slide$get(), "//p:spTree"), template_node)
  }
  slide$save()
  x
}



#' @export
slide_info <- function( x, layout, master ){

  master_info <- x$slideLayouts$get_master()$get_data()
  master_info$master_file <- basename( master_info$filename )
  coordinates <- inner_join(x$coordinates, master_info, by = 'master_file')
  line_select <- coordinates$name %in% layout & coordinates$master_name %in% master
  coordinates[line_select, ]

}


