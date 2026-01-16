library(officer)

x <- read_pptx()
layout_properties(x = x, layout = "Title Slide", master = "Office Theme")
layout_properties(x = x, master = "Office Theme")
layout_properties(x = x, layout = "Two Content")
layout_properties(x = x)
