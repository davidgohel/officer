wml_str <- function(str){
  paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
          "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
          str,
          "</w:document>"
  )
}

pml_str <- function(str){
  paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
          sprintf("<a:document%s>", officer:::pml_ns),
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

