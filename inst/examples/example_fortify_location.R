library(officer)

doc <- read_pptx()
doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
fortify_location(ph_location_fullsize(), doc)
