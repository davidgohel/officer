context("pptx selection cases")

library(xml2)
library(magrittr)

test_that("on_slide", {
  x <- read_pptx()
  expect_error(slide_summary(x), "presentation contains no slide")
  expect_error(remove_slide(x), "presentation contains no slide to delete")
  expect_error(on_slide(x, index = 1), "presentation contains no slide")

  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  x <- ph_with_text(x, type = "title", str = "my title")
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  x <- ph_with_text(x, type = "title", str = "my title 2")

  expect_error(on_slide(x, index = 3), "unvalid index 3")
  expect_error(remove_slide(x, index = 3), "unvalid index 3")
  expect_error(slide_summary(x, index = 3), "unvalid index 3")

})

unlink("*.pptx")

