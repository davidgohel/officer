library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Comparison")
plot_layout_properties(doc, "Comparison")

doc <- ph_with(doc, "The Title", location = ph_location_id(id = 2)) # title
doc <- ph_with(doc, "Left Header", location = ph_location_id(id = 3)) # left header
doc <- ph_with(doc, "Left Content", location = ph_location_id(id = 4)) # left content
doc <- ph_with(doc, "The Footer", location = ph_location_id(id = 8)) # footer

file <- tempfile(fileext = ".pptx")
print(doc, file)

## file.show(file) # may not work on your system
