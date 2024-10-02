test_that("get_layout works as expected", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  # all layouts unique
  x <- read_pptx()

  expect_error(get_layout(x, "Title Slide", "xxx"), 'master "xxx" does not exist')
  expect_error(get_layout(x, "xxx", "Office Theme"), 'Layout "xxx" does not exist in master "Office Theme"')

  expect_error(get_layout(x, 0), "Layout index out of bounds.")
  expect_error(get_layout(x, 10), "Layout index out of bounds.")
  expect_error(get_layout(x, Inf), "Layout index out of bounds.")

  expect_no_error(l1 <- get_layout(x, "Title Slide"))
  expect_s3_class(l1, "layout_info")
  expect_no_error(l2 <- get_layout(x, "Title Slide", "Office Theme"))
  expect_s3_class(l2, "layout_info")

  expect_no_error(la <- get_layout(x, 1))
  expect_equal(la$index, 1)

  # same layout in several masters
  file <- test_path("docs_dir", "test-three-identical-masters.pptx")
  x <- read_pptx(file)

  expect_error(get_layout(x, "xxx", NULL), 'Layout "xxx" does not exist')
  expect_error(get_layout(x, "xxx", "Master_1"), 'Layout "xxx" does not exist in master "Master_1"')
  expect_error(get_layout(x, "Title Slide"), "Layout exists in more than one master")
  expect_error(get_layout(x, "Title Slide", "xxx"), 'master "xxx" does not exist')

  expect_no_error(la <- get_layout(x, "Title Slide", "Master_1"))
  expect_s3_class(la, "layout_info")

  expect_no_error(la <- get_layout(x, 1))
  expect_equal(la$index, 1)
})


test_that("incorrect inputs are detected", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()
  layout <- "Comparison"

  expect_error(get_layout("no rpptx object", layout), "Incorrect input for `x`")

  # incorrect layout arg input
  error_msg <- "`layout` must be <numeric> or <character>"
  expect_error(get_layout(x, NA), error_msg)
  expect_error(get_layout(x, NULL), error_msg)
  expect_error(get_layout(x, mtcars), error_msg)

  # out of bound index
  error_msg <- "Layout index out of bounds"
  expect_error(get_layout(x, 0), error_msg)
  expect_error(get_layout(x, -1), error_msg)
  expect_error(get_layout(x, 8), error_msg)
  expect_error(get_layout(x, Inf), error_msg)

  # non existing layout
  expect_error(get_layout(x, "xxx"), 'Layout "xxx" does not exist.')
  expect_error(get_layout(x, ""), 'Layout "" does not exist.')

  # not exactly one layout
  error_msg <- "`layout` is not length 1"
  expect_error(get_layout(x, c("a", "b")), error_msg)
  expect_error(get_layout(x, 1:2), error_msg)
  expect_error(get_layout(x, integer()), error_msg)
  expect_error(get_layout(x, character()), error_msg)

  # layout not unique
  file <- test_path("docs_dir", "test-three-identical-masters.pptx")
  x <- read_pptx(file)
  expect_error(get_layout(x, "Title Slide"), "Layout exists in more than one master")
})


test_that("<layout_info> prints correctly", {
  x <- read_pptx()
  layout <- "Comparison"
  l <- get_layout(x, layout)
  out <- capture.output(print(l))
  expect_equal(length(out), length(l))
})


test_that("get layout from slide", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()

  # fails if no slides exist
  expect_error(get_slide_layout(x, 0), "Presentation does not have any slides yet")
  expect_error(get_layout_for_current_slide(x), "Presentation does not have any slides yet")

  # detect correct slide layout
  layout <- "Comparison"
  x <- add_slide(x, layout)

  expect_error(get_slide_layout(x, 0), "Slide index 0 is out of range")
  expect_error(get_slide_layout(x, 2), "Slide index 2 is out of range")

  error_msg <- "`slide_idx` must be <numeric>"
  expect_error(get_slide_layout(x, "1"), error_msg)
  expect_error(get_slide_layout(x, NA), error_msg)
  expect_error(get_slide_layout(x, NULL), error_msg)

  expect_no_error(get_slide_layout(x, 1))
  expect_no_error(get_layout_for_current_slide(x))

  la_slide <- get_slide_layout(x, 1)
  la_current <- get_layout_for_current_slide(x)
  la_reference <- get_layout(x, layout)
  expect_identical(la_current, la_reference)
  expect_identical(la_slide, la_reference)

  layout <- "Title Slide"
  x <- add_slide(x, layout)
  la_slide <- get_slide_layout(x, 2)
  la_current <- get_layout_for_current_slide(x)
  la_reference <- get_layout(x, layout)
  expect_identical(la_current, la_reference)
  expect_identical(la_slide, la_reference)
})
