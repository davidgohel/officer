context("read_pptx")

library(utils)
library(xml2)

test_that("defaul template", {
  x <- read_pptx()
  expect_equal(length( x ), 0)
  expect_true(file.exists(x$package_dir))
})

test_that("console printing", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  expect_output(print(x), "pptx document with 1 ")
})

test_that("check extention and print document", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  print(x, target = "print.pptx")
  expect_true( file.exists("print.pptx") )

  expect_error(print(x, target = "print.pptxxx"))
})

test_that("check slide_summary", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  sm <- slide_summary(x)
  expect_true(nrow(sm) == 1)
  expect_true(ncol(sm) == 7)
})

test_that("check template", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  print(x, target = "template.pptx")

  expect_silent(x <- read_pptx(path = "template.pptx"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world again")
  print(x, target = "example.pptx")

  expect_equal(length(x), 2)
})

unlink("*.pptx")
unlink("*.emf")
