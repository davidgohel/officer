library(officer)

doc <- read_docx()
doc <- body_add_par(doc, value = "blah blah blah", style = "Normal")

z <- body_append_start_context(doc)

for (i in seq_len(50)) {
  write_elements_to_context(
    context = z,
    fpar(
      "Hello World, ",
      i,
      fp_p = fp_par(word_style = "heading 1")
    ),
    fpar(run_pagebreak())
  )
}
doc <- body_append_stop_context(z)


print(doc, target = tempfile(fileext = ".docx"))
