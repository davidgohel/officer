context("doc summary")

test_that("docx summary", {
  example_docx <- system.file(package = "officer", "doc_examples/example.docx")

  doc <- read_docx(path = example_docx)
  doc_data <- docx_summary(doc)
  table_data <- subset(doc_data, content_type %in% "table cell" & is_header)
  expect_equal( table_data$text, c("Petals", "Internode", "Sepal", "Bract") )
})

test_that("complex docx table", {

  doc <- read_docx(path = "docs_dir/table-complex.docx" )
  doc_data <- docx_summary(doc)
  table_data <- doc_data[doc_data$content_type == "table cell", ]
  table_data <- table_data[order(table_data$row_id, table_data$cell_id),]

  first_row <- table_data[table_data$row_id %in% 1,]

  expect_equal( first_row$text, c("Column head", "column head", "column head", "column head",
                                  NA, NA, "x", "y") )
  expect_equal( first_row$col_span, c(1, 1, 1, 3, 0, 0, 1, 1) )
  expect_equal( first_row$row_span, c(2, 2, 2, 1, 1, 1, 2, 2) )
  expect_true( all( table_data$text[table_data$col_span < 1 | table_data$row_span < 1] %in% NA_character_ ) )
})




test_that("pptx summary", {
  example_pptx <- system.file(package = "officer", "doc_examples/example.pptx")
  doc <- read_pptx(path = example_pptx)
  doc_data <- pptx_summary(doc)
  table_data <- subset(doc_data, content_type %in% "table cell" & row_id == 1 & slide_id == 1)
  expect_equal( table_data$text, c("Header 1 ", "Header 2", "Header 3") )

  table_data <- subset(doc_data, content_type %in% "table cell" & row_id == 4 & slide_id == 1)
  expect_equal( table_data$text, c("B", "9.0", "Salut") )

  example_pptx <- "docs_dir/table-complex.pptx"
  doc <- read_pptx(path = example_pptx)
  doc_data <- pptx_summary(doc)
  table_data <- subset(doc_data, content_type %in% "table cell" & cell_id == 3 & slide_id == 1)
  expect_equal( table_data$text,
                c("Header 3", NA, "blah blah blah", "Salut", "Hello", "sisi", NA, NA, NA, NA) )
  expect_equal( table_data$row_span,
                c(1, 1, 1, 1, 1, 4, 0, 0, 0, 1) )
  expect_equal( table_data$col_span,
                c(1, 0, 1, 1, 1, 1, 1, 1, 1, 0) )

  slide2_data <- subset(doc_data, content_type %in% "paragraph" & slide_id == 2)
  expect_equal( slide2_data$text, c("coco", "line of text", "blah blah blah") )
})

