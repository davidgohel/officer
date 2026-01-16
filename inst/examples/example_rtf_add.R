library(officer)

def_text <- fp_text_lite(color = "#006699", bold = TRUE)
center_par <- fp_par(text.align = "center", padding = 3)

doc <- rtf_doc(
  normal_par = fp_par(line_spacing = 1.4, padding = 3)
)

doc <- rtf_add(
  x = doc,
  value = fpar(
    ftext("how are you?", prop = def_text),
    fp_p = fp_par(text.align = "center")
  )
)

a_paragraph <- fpar(
  ftext("Here is a date: ", prop = def_text),
  run_word_field(field = "Date \\@ \"MMMM d yyyy\""),
  fp_p = center_par
)
doc <- rtf_add(
  x = doc,
  value = block_list(
    a_paragraph,
    a_paragraph,
    a_paragraph
  )
)

if (require("ggplot2")) {
  gg <- gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))
  doc <- rtf_add(doc, gg, width = 3, height = 4, ppr = center_par)
}
anyplot <- plot_instr(code = {
  barplot(1:5, col = 2:6)
})
doc <- rtf_add(doc, anyplot, width = 5, height = 4, ppr = center_par)

print(doc, target = tempfile(fileext = ".rtf"))
