context("border formatting properties")

library(xml2)
library(magrittr)

test_that("fp_border", {
  expect_error( fp_border(width = -5), "width must be a positive integer scalar" )
  expect_error( fp_border(color = "glop"), "color must be a valid color" )
  expect_error( fp_border(style = "glop"), "style must be one of" )
  x <- fp_border(color = "red", style = "dashed", width = 5)
  x <- update(x, color = "yellow", style = "solid", width = 2)
  expect_equal(x$color, "yellow")
  expect_equal(x$style, "solid")
  expect_equal(x$width, 2)


})

unlink("*.pptx")

