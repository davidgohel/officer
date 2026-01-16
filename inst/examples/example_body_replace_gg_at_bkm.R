library(officer)

if (require("ggplot2")) {
  gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))

  doc <- read_docx()
  doc <- body_add_par(doc, "insert_plot_here")
  doc <- body_bookmark(doc, "plot")
  doc <- body_replace_gg_at_bkm(doc, bookmark = "plot", value = gg_plot)
  print(doc, target = tempfile(fileext = ".docx"))
}
doc <- read_docx()
doc <- body_add_par(doc, "insert_plot_here")
doc <- body_bookmark(doc, "plot")
if (capabilities(what = "png")) {
  doc <- body_replace_plot_at_bkm(
    doc,
    bookmark = "plot",
    value = plot_instr(
      code = {
        barplot(1:5, col = 2:6)
      }
    )
  )
}
print(doc, target = tempfile(fileext = ".docx"))
