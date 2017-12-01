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

test_that("names stay as is", {
  df <- data.frame("hello coco"=c(1, 2), value=c("одно значение", "еще значение"),
                   check.names = FALSE, stringsAsFactors = FALSE)
  x <- read_docx()
  x <- body_add_table(x, value = df, style = "table_template")
  node <- x$doc_obj$get_at_cursor()
  first_text <- xml_find_first(node, xpath = "w:tr[w:trPr/w:tblHeader]/w:tc/w:p/w:r/w:t")
  first_text <- xml_text(first_text)
  expect_equal(first_text, "hello coco")
})

test_that("tibles are casted to data.frame", {
  testthat::skip_if_not(requireNamespace("tibble"))
  df <- tibble::tibble("hello coco"=c(1, 2), value=c("a", "b"))
  x <- read_docx()
  x <- body_add_table(x, value = df, style = "table_template")
  node <- x$doc_obj$get_at_cursor()
  first_text <- xml_find_first(node, xpath = "w:tr[w:trPr/w:tblHeader]/w:tc/w:p/w:r/w:t")
  first_text <- xml_text(first_text)
  expect_equal(first_text, "hello coco")
})

