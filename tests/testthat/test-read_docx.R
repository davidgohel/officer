context("read_docx")

library(utils)
library(xml2)

test_that("defaul template", {
  x <- read_docx()
  expect_equal(length( x ), 2)
  expect_true(file.exists(x$doc_obj$package_dirname()))
})

test_that("console printing", {
  x <- read_docx()
  expect_output(print(x), "docx document with")
})

test_that("check extention and print document", {
  x <- read_docx()
  print(x, target = "print.docx")
  expect_true( file.exists("print.docx") )

  expect_error(print(x, target = "print.docxxx"))
})

