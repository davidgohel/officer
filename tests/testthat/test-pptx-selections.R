test_that("check slide selection", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 1", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 2", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 3", location = ph_location_type(type = "body"))

  x <- on_slide(x, index = 1)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 1")

  x <- on_slide(x, index = 2)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 2")

  x <- on_slide(x, index = 3)
  sm <- slide_summary(x)
  expect_equal(sm[1, ]$text, "Hello world 3")
})


test_that("check errors", {
  x <- read_pptx()
  expect_error(slide_summary(x), "presentation contains no slide")
  expect_error(remove_slide(x), "presentation contains no slide to delete")
  expect_error(on_slide(x, index = 1), "presentation contains no slide")

  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  x <- ph_with(x, "my title", location = ph_location_type(type = "title"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  x <- ph_with(x, "my title 2", location = ph_location_type(type = "title"))

  expect_error(on_slide(x, index = 3), "unvalid index 3")
  expect_error(remove_slide(x, index = 3), "unvalid index 3")
  expect_error(slide_summary(x, index = 3), "unvalid index 3")
})


test_that("get shape id", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- ph_with(doc, "hello", location = ph_location_type(type = "body"))
  file <- print(doc, target = tempfile(fileext = ".pptx"))
  doc <- read_pptx(file)
  expect_equal(officer:::get_shape_id(doc, type = "body", id = 1), "2")
  expect_equal(
    officer:::get_shape_id(doc, ph_label = "Content Placeholder 2", id = 1),
    "2"
  )
  expect_error(officer:::get_shape_id(doc, type = "body", id = 4))
})


unlink("*.pptx")


test_that("ensure_slide_index_exists ", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()

  expect_error(ensure_slide_index_exists(x, 0), "Slide index 0 is out of range")
  expect_error(ensure_slide_index_exists(x, 2), "Slide index 2 is out of range")

  x <- add_slide(x, "Comparison")
  expect_no_error(ensure_slide_index_exists(x, 1))
  expect_error(ensure_slide_index_exists(x, 0), "Slide index 0 is out of range")
  expect_error(ensure_slide_index_exists(x, 2), "Slide index 2 is out of range")

  error_msg <- "`slide_idx` must be <numeric>"
  expect_error(ensure_slide_index_exists(x, "1"), error_msg)
  expect_error(ensure_slide_index_exists(x, NA), error_msg)
  expect_error(ensure_slide_index_exists(x, NULL), error_msg)
})
