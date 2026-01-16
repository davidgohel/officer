library(officer)

example_docx <- system.file(
  package = "officer",
  "doc_examples/example.docx"
)
doc <- read_docx(example_docx)

docx_summary(doc)

docx_summary(doc, detailed = TRUE)
