context("doc summary")

test_that("docx summary", {
  example_docx <- system.file(package = "officer", "doc_examples/example.docx")

  doc <- read_docx(path = example_docx)
  doc_data <- docx_summary(doc)
  table_data <- dplyr::filter(doc_data, content_type == "table cell", is_header )

  expect_equal( table_data$text, c("Petals", "Internode", "Sepal", "Bract") )
})

test_that("pptx summary", {
  example_pptx <- system.file(package = "officer", "doc_examples/example.pptx")
  doc <- read_pptx(path = example_pptx)
  doc_data <- pptx_summary(doc)
  table_data <- dplyr::filter(doc_data, content_type == "table cell", row_id == 1, slide_id == 1)
  expect_equal( table_data$text, c("Header 1 ", "Header 2", "Header 3") )

  table_data <- dplyr::filter(doc_data, content_type == "table cell", row_id == 4, slide_id == 1)
  expect_equal( table_data$text, c("B", "9.0", "Salut") )
})

