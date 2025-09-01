test_that("body_import_docx messages", {
  tf <- tempfile(fileext = ".docx")

  x <- read_docx()
  testthat::expect_warning(
    x <- body_import_docx(
      x = x,
      src = system.file(
        package = "officer", "doc_examples", "example.docx"
      )
    ),
    "Style(s) mapping(s) for 'paragraphs' are missing in the body of the document:",
    fixed = TRUE
  )

  x <- read_docx()
  testthat::expect_no_message(
    x <- body_import_docx(
      x = x,
      src = system.file(
        package = "officer", "doc_examples", "example.docx"
      ),
      par_style_mapping = list(
        "Normal" = c("List Paragraph")
      ),
      tbl_style_mapping = list(
        "Normal Table" = "Light Shading"
      )
    )
  )
  docx_summary(x)
})

test_that("body_import_docx copy the content", {
  tf <- tempfile(fileext = ".docx")
  fi <- system.file(
    package = "officer", "doc_examples", "example.docx"
  )
  x <- read_docx()
  x <- body_import_docx(
    x = x,
    src = fi,
    par_style_mapping = list(
      "Normal" = c("List Paragraph")
    ),
    tbl_style_mapping = list(
      "Normal Table" = "Light Shading"
    )
  )
  print(x, target = tf)

  old_doc <- read_docx(fi)
  old_doc_head <- docx_summary(old_doc)

  new_doc <- read_docx(tf)
  new_doc_head <- docx_summary(new_doc)

  expect_equal(
    head(new_doc_head)[c("doc_index", "text")],
    head(old_doc_head)[c("doc_index", "text")]
  )
  expect_equal(
    head(new_doc_head)$style_name,
    c("heading 1", NA, "heading 1", "Normal", "Normal", "Normal")
  )

})

