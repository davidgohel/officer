library(officer)

doc <- read_docx()
doc <- body_add_par(doc, "a paragraph to replace", style = "centered")
doc <- body_bookmark(doc, "text_to_replace")
doc <- body_replace_text_at_bkm(doc, "text_to_replace", "new text")


# demo usage of bookmark and images ----
template <- system.file(package = "officer", "doc_examples/example.docx")

img.file <- file.path(R.home("doc"), "html", "logo.jpg")

doc <- read_docx(path = template)
doc <- headers_replace_img_at_bkm(
  x = doc,
  bookmark = "bmk_header",
  value = external_img(src = img.file, width = .53, height = .7)
)
doc <- footers_replace_img_at_bkm(
  x = doc,
  bookmark = "bmk_footer",
  value = external_img(src = img.file, width = .53, height = .7)
)
print(doc, target = tempfile(fileext = ".docx"))
