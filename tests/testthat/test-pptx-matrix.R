test_that("as.matrix.rpptx", {
  example_pptx <- test_path("docs_dir/table-complex.pptx")
  doc <- read_pptx(path = example_pptx)
  # Single table, without fill
  expect_equal(
    as.matrix(doc),
    matrix(
      c("Header 1 A", NA, "B", "B", "C", "", "", "", "", 
        "tutu", "Header 2", "12.23blah blah", "1.23", "9.0", "6", "", 
        "", "", "", NA, "Header 3", NA, "blah blah blah", "Salut", "Hello", 
        "sisi", NA, NA, NA, NA),
      nrow=10, ncol=3
    )
  )
  # Single table, with fill
  expect_equal(
    as.matrix(doc, span="fill"),
    matrix(
      c("Header 1 A", "Header 1 A", "B", "B", "C", "", "", 
        "", "", "tutu", "Header 2", "12.23blah blah", "1.23", "9.0", 
        "6", "", "", "", "", "tutu", "Header 3", "12.23blah blah", "blah blah blah", 
        "Salut", "Hello", "sisi", "sisi", "sisi", "sisi", "tutu"),
      nrow=10, ncol=3
    )
  )
  # Single table, specify slide_id and id gets the first table
  expect_equal(
    as.matrix(doc, slide_id=1, id=18),
    as.matrix(doc)
  )
  # All tables, without fill
  expect_equal(
    as.matrix(doc, slide_id=NULL),
    list("1"=list("18"=matrix(
      c("Header 1 A", NA, "B", "B", "C", "", "", "", "", 
        "tutu", "Header 2", "12.23blah blah", "1.23", "9.0", "6", "", 
        "", "", "", NA, "Header 3", NA, "blah blah blah", "Salut", "Hello", 
        "sisi", NA, NA, NA, NA),
      nrow=10, ncol=3
    )))
  )
})
