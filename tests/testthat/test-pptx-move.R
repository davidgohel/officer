test_that("check errors", {
  x <- read_pptx()
  expect_error(move_slide(x, index = 2, to = 1), "presentation contains no slide")

  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 1", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 2", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 3", location = ph_location_type(type = "body"))

  x <- on_slide(x, index = 1)
  sm <- slide_summary(x)
  expect_equal(sm[1,]$text, "Hello world 1")

  expect_error(move_slide(x, to = 4))
  expect_error(move_slide(x, index = 5, to = 4))
  x <- move_slide(x, to = 3)
  x <- on_slide(x, index = 3)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 1")
  x <- remove_slide(x, index = 3)
  x <- move_slide(x, index = 2, to = 1)
  x <- on_slide(x, index = 1)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 3")
})


unlink("*.pptx")

