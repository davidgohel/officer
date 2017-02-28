context("table in docx")

test_that("body_add_table", {
  x <- read_docx()
  x <- body_add_table(x, value = iris, style = "table_template")
  node <- x$doc_obj$get_at_cursor()
  expect_equal( xml_name(node), "tbl" )
})

test_that("wml structure", {
  x <- read_docx()
  x <- body_add_table(x, value = iris, style = "table_template")
  node <- x$doc_obj$get_at_cursor()
  expect_length(xml_find_all(node, xpath = "w:tr[w:trPr/w:tblHeader]"), 1)
  expect_length(xml_find_all(node, xpath = "w:tr[not(w:trPr/w:tblHeader)]"), nrow(iris))
  expect_length(xml_find_all(node, xpath = "w:tr[not(w:trPr/w:tblHeader)]/w:tc/w:p/w:r/w:t"), ncol(iris)*nrow(iris))
})

