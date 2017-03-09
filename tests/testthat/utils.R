wml_str <- function(str){
  paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
          officer:::wml_with_ns("w:document"),
          str,
          "</w:document>"
  )
}

pml_str <- function(str){
  paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
          officer:::pml_with_ns("a:document"),
          str,
          "</a:document>"
  )
}

has_css_color <- function(x, atname, color){
  css <- format(x, type = "html")
  reg <- sprintf("%s:%s", atname, color)
  grepl(reg, css)
}

has_css_attr <- function(x, atname, value){
  css <- format(x, type = "html")
  reg <- sprintf("%s:%s", atname, value)
  grepl(reg, css)
}



