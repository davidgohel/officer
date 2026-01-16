library(officer)

doc <- read_docx()

img.file <- file.path(R.home("doc"), "html", "logo.jpg")
if (file.exists(img.file)) {
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39)

  # Set the unit in which the width and height arguments are expressed
  doc <- body_add_img(
    x = doc,
    src = img.file,
    height = 2.69,
    width = 3.53,
    unit = "cm"
  )
}

print(doc, target = tempfile(fileext = ".docx"))
