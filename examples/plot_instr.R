# plot_instr demo ----

anyplot <- plot_instr(code = {
  barplot(1:5, col = 2:6)
  })

doc <- read_pptx()
doc <- add_slide(doc)
doc <- ph_with(
  doc, anyplot,
  location = ph_location_fullsize(),
  bg = "#00000066", pointsize = 12)
print(doc, target = tempfile(fileext = ".pptx"))
