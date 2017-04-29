context("checks selections for pptx")

test_that("check slide selection", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 1")
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 2")
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 3")

  x <-  x %>% on_slide(index = 1)
  sm <- slide_summary(x)
  expect_equal(sm[1,]$text, "Hello world 1")

  x <-  x %>% on_slide(index = 2)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 2")

  x <-  x %>% on_slide(index = 3)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 3")

})

test_that("check errors", {
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

test_that("get shape id", {

  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "hello")
  expect_equal(officer:::get_shape_id(doc, type = "body", id_chr = "2"), 1)
  expect_equal(officer:::get_shape_id(doc, type = "body"), 1)
  expect_equal(officer:::get_shape_id(doc, id_chr = "2"), 1)
  expect_error(officer:::get_shape_id(doc, type = "body", id_chr = "1") )
})


unlink("*.pptx")

