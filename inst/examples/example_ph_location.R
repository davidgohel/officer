library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(
  doc,
  "Hello world",
  location = ph_location(width = 4, height = 3, newlabel = "hello")
)
print(doc, target = tempfile(fileext = ".pptx"))

# Set geometry and outline
doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
loc <- ph_location(
  left = 1,
  top = 1,
  width = 4,
  height = 3,
  bg = "steelblue",
  ln = sp_line(color = "red", lwd = 2.5),
  geom = "trapezoid"
)
doc <- ph_with(doc, "", loc = loc)
print(doc, target = tempfile(fileext = ".pptx"))
