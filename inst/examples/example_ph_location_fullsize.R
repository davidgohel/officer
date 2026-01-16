library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, "Hello world", location = ph_location_fullsize())
print(doc, target = tempfile(fileext = ".pptx"))
