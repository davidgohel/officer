context("miscellaneous checks for xlsx")

test_that("console printing", {
  x <- read_xlsx()
  expect_output(print(x), regexp = "^xlsx document with 1 sheet")
})

test_that("check extention and print document", {
  x <- read_xlsx()
  print(x, target = "print.xlsx")
  expect_true( file.exists("print.xlsx") )
  expect_error(print(x, target = "print.xlsxxxx"))
})


unlink("*.xlsx")

