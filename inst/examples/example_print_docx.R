library(officer)

# This example demonstrates how to create
# an small document -----

## Create a new Word document
doc <- read_docx()
doc <- body_add_par(doc, "hello world")
## Save the document
output_file <- print(doc, target = tempfile(fileext = ".docx"))

# preview mode: save to temp file and open locally ----
## Not run:
# print(doc, preview = TRUE)
