library(officer)

# create a template ----
doc <- read_docx()
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "Hello text to replace")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "Hello text to replace")
doc <- body_add_par(doc, "blah blah blah")
template_file <- print(
  x = doc,
  target = tempfile(fileext = ".docx")
)

# replace all pars containing "to replace" ----
doc <- read_docx(path = template_file)
while (cursor_reach_test(doc, "to replace")) {
  doc <- cursor_reach(doc, "to replace")

  doc <- body_add_fpar(
    x = doc,
    pos = "on",
    value = fpar(
      "Here is a link: ",
      hyperlink_ftext(
        text = "yopyop",
        href = "https://cran.r-project.org/"
      )
    )
  )
}

doc <- cursor_end(doc)
doc <- body_add_par(doc, "Yap yap yap yap...")

result_file <- print(
  x = doc,
  target = tempfile(fileext = ".docx")
)

# cursor_bookmark ----

doc <- read_docx()
doc <- body_add_par(doc, "centered text", style = "centered")
doc <- body_bookmark(doc, "text_to_replace")
doc <- body_add_par(doc, "A title", style = "heading 1")
doc <- body_add_par(doc, "Hello world!", style = "Normal")
doc <- cursor_bookmark(doc, "text_to_replace")
doc <- body_add_table(doc, value = iris, style = "table_template")

print(doc, target = tempfile(fileext = ".docx"))
