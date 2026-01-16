library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- add_slide(doc, "Title and Content")
doc <- add_slide(doc, "Title and Content")
doc <- on_slide(doc, index = 1)
doc <- ph_with(
  x = doc,
  "First title",
  location = ph_location_type(type = "title")
)
doc <- on_slide(doc, index = 3)
doc <- ph_with(
  x = doc,
  "Third title",
  location = ph_location_type(type = "title")
)

file <- tempfile(fileext = ".pptx")
print(doc, target = file)
