library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, "Title", location = ph_location_type(type = "title"))
doc <- ph_with(
  doc,
  "Hello world",
  location = ph_location_template(top = 4, type = "title")
)
print(doc, target = tempfile(fileext = ".pptx"))
