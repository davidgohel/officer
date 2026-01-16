library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, "Hello left", location = ph_location_left())
doc <- ph_with(doc, "Hello right", location = ph_location_right())
print(doc, target = tempfile(fileext = ".pptx"))
