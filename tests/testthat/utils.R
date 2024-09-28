wml_str <- function(str){
  paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
          "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
          str,
          "</w:document>"
  )
}

pml_str <- function(str){
  paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
          "<a:document xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
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
