# rename function 
layout_duplicate_rename <- function(x) {
  
  for (slide_layout in split(x$slideLayouts$get_xfrm_data(),factor(x$slideLayouts$get_xfrm_data()$file))) {
    
    for(duplicated_layout in unique(slide_layout$ph_label[duplicated(slide_layout$ph_label)])) {
      
      duplicated_ids <- as.numeric(
        slide_layout[
          slide_layout$ph_label %in% slide_layout$ph_label[duplicated(slide_layout$ph_label)],
        ]$id
      )
      
      rename_ids <- 
        setdiff(duplicated_ids, min(duplicated_ids))
      
      layout_file <- 
        list.files(x$package_dir, paste0(slide_layout$file[1], "$"), recursive = T, full.names = T)
      
      layout_xml <- 
        xml2::read_xml(layout_file)
      
      for (rename_index in 1:length(rename_ids)) {
        
        layout_delete_nodes <-
          xml2::xml_find_all(layout_xml, sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", rename_ids[rename_index]))
        
        xml2::xml_set_attr(
          xml2::xml_find_all(
            layout_xml, 
            sprintf("//p:cNvPr[@id='%s']", rename_ids[rename_index])
          ),
          "name", paste0(duplicated_layout, " ", rename_index)
        )
        
      }
      
      xml2::write_xml(layout_xml, file =layout_file)
      
    }
    
  }
}
# delete function 
layout_duplicate_delete <- function(x, keep_max_id = TRUE) {
  
  for (slide_layout in split(x$slideLayouts$get_xfrm_data(),factor(x$slideLayouts$get_xfrm_data()$file))) {
    
    for(duplicated_layout in unique(slide_layout$ph_label[duplicated(slide_layout$ph_label)])) {
      
      duplicated_ids <- as.numeric(
        slide_layout[
          slide_layout$ph_label %in% slide_layout$ph_label[duplicated(slide_layout$ph_label)],
        ]$id
      )
      
      delete_ids <- 
        setdiff(duplicated_ids, ifelse(keep_max_id, max(duplicated_ids), min(duplicated_ids)))
      
      layout_file <- 
        list.files(x$package_dir, paste0(slide_layout$file[1], "$"), recursive = T, full.names = T)
      
      layout_xml <- 
        xml2::read_xml(layout_file)
      
      for (delete_id in delete_ids) {
        
        layout_delete_node <-
          xml2::xml_find_all(layout_xml, sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", delete_id))
        
        xml2::xml_remove(layout_delete_node)
        
      }
      
      xml2::write_xml(layout_xml, file =layout_file)
      
    }
    
  }
}