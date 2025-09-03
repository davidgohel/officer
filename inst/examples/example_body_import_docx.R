library(officer)

# example file from the package
file_input <- system.file(
  package = "officer",
  "doc_examples/example.docx"
)

# create a new rdocx document
x <- read_docx()

# import content from file_input
x <- body_import_docx(
  x = x,
  src = file_input,
  # style mapping for paragraphs and tables
  par_style_mapping = list(
    "Normal" = c("List Paragraph")
  ),
  tbl_style_mapping = list(
    "Normal Table" = "Light Shading"
  )
)

# Create temporary file
tf <- tempfile(fileext = ".docx")
# write to file
print(x, target = tf)
