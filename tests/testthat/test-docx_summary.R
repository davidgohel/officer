context("docx summary")

test_that("docx summary", {

  doc <- read_docx(path = "template/template.docx")
  doc_data <- docx_summary(doc)
  table_data <- doc_data[[1, "table_data"]]

  expect_equal( table_data$txt[table_data$row_id %in% 1], c("G", "H", "I") )

  expect_equal( doc_data[2, ]$txt, "Text 1")
  expect_equal( doc_data[2, ]$style_name, "table title")
})

