## RTF example with sections ----

library(officer)

quick_section_header <- function(label) {
  block_list(
    fpar(
      ftext(label, fp_text_lite(bold = TRUE, color = "#006699"))
    )
  )
}

quick_section_footer <- function(label) {
  block_list(
    fpar(
      "Page ",
      run_word_field(field = "PAGE  \\* MERGEFORMAT")
    )
  )
}

quick_hello_world <- function(doc) {
  rtf_add(
    doc,
    fpar(
      ftext(
        "Hello World"
      ),
      fp_p = fp_par(text.align = "left")
    )
  )
}

three_cols_section <- block_section(
  prop_section(
    type = "continuous",
    section_columns = section_columns(widths = c(1.7, 1.7, 1.7), space = 0.25),
    header_default = quick_section_header("Three columns section"),
    footer_default = quick_section_footer("Three columns section")
  )
)

doc <- rtf_doc(
  def_sec = prop_section(
    header_default = quick_section_header("Default section"),
    footer_default = quick_section_footer("Default section")
  ),
  normal_par = fp_par(padding = 3)
)

doc <- quick_hello_world(doc)

if (require("ggplot2")) {
  gg_iris <- ggplot(iris, aes(Sepal.Length, Petal.Length, colour = Species)) +
    geom_point() +
    theme_minimal()
  doc <- rtf_add(doc, gg_iris, width = 4, height = 3)
}

doc <- rtf_add(doc, three_cols_section)

titles_list <- block_list(
  fpar(ftext("Left Title"), fp_p = fp_par(text.align = "left")),
  fpar(
    run_columnbreak(),
    ftext("Centered Title"),
    fp_p = fp_par(text.align = "center")
  ),
  fpar(
    run_columnbreak(),
    ftext("Right Title"),
    fp_p = fp_par(text.align = "right")
  )
)
doc <- rtf_add(doc, titles_list)
doc <- rtf_add(doc, block_section(prop_section()))
doc <- rtf_add(doc, fpar(run_linebreak()))
doc <- quick_hello_world(doc)

landscape_section <- block_section(prop_section(
  type = "nextPage",
  page_size = page_size(orient = "landscape"),
  header_default = quick_section_header("Landscape section"),
  footer_default = quick_section_footer("Landscape section")
))
doc <- rtf_add(doc, landscape_section)

doc <- quick_hello_world(doc)

doc <- rtf_add(
  doc,
  block_section(
    prop_section(
      type = "nextPage",
      header_default = quick_section_header("Final section"),
      footer_default = quick_section_footer("Final section")
    )
  )
)

doc <- quick_hello_world(doc)

print(doc, target = tempfile(fileext = ".rtf"))
