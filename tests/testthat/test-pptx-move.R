context("checks slide moving")

test_that("check errors", {
  x <- read_pptx()
  expect_error(move_slide(x, index = 2, to = 1), "presentation contains no slide")

  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 1")
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 2")
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 3")

  x <- x %>% on_slide(index = 1)
  sm <- slide_summary(x)
  expect_equal(sm[1,]$text, "Hello world 1")

  x <- x %>% move_slide(index = 1, to = 3) %>%
    on_slide(index = 3)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 1")
  x <- remove_slide(x, index = 3)
  x <- x %>% move_slide(index = 2, to = 1) %>%
    on_slide(index = 1)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 3")
})


unlink("*.pptx")

