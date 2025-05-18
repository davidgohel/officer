test_that("trailing file index extraction / sorting", {
  files <- c(
    "slideLayout1.xml",
    "slideLayout11.xml",
    "slideLayout2.xml",
    "slideLayout10.xml"
  )

  expect_equal(get_file_index(files), c(1, 11, 2, 10))

  expect_equal(
    sort_vec_by_index(files),
    c(
      "slideLayout1.xml",
      "slideLayout2.xml",
      "slideLayout10.xml",
      "slideLayout11.xml"
    )
  )

  df <- data.frame(file1 = files, file2 = rev(files))
  a <- sort_dataframe_by_index(df, "file1")
  b <- sort_dataframe_by_index(df, "file2")
  expect_true(all(a == rev(b)))
})


test_that("misc", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  df <- df_rename(mtcars, c("mpg", "cyl"), c("A", "B"))
  expect_true(all(names(df)[1:2] == c("A", "B")))

  expect_true(is_integerish(1))
  expect_true(is_integerish(1:3))
  expect_true(is_integerish(1.0))
  expect_true(is_integerish(c(1.0, 2.0)))
  expect_false(is_integerish(1.00001))
  expect_false(is_integerish(c(1.1, 2.0)))
  expect_false(is_integerish(TRUE))
  expect_false(is_integerish(FALSE))
  expect_error(
    expect_warning(stop_if_not_integerish(LETTERS)),
    regex = "Expected integerish values but got <character>"
  )
  expect_error(
    stop_if_not_integerish(c(1.1, 1.2)),
    regex = "Expected integerish values but got <numeric>"
  )
})


test_that("stop_if_not_in_slide_range", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()
  expect_error(
    regex = "Presentation has no slides",
    stop_if_not_in_slide_range(x, 1)
  )

  x <- add_slide(x, "Title and Content")
  x <- add_slide(x, "Title and Content")
  expect_no_error(stop_if_not_in_slide_range(x, 1:2))
  expect_error(
    regex = "1 index outside slide range",
    stop_if_not_in_slide_range(x, -1)
  )
  expect_error(
    regex = "2 indexes of `my_arg` outside slide range",
    stop_if_not_in_slide_range(x, 3:4, arg = "my_arg")
  )
  expect_error(
    regex = "Expected integerish values but got <numeric>",
    stop_if_not_in_slide_range(x, 3.1)
  )

  foo <- function() {
    stop_if_not_in_slide_range(x, 3:4)
  }
  error_text <- tryCatch(foo(), error = paste)
  grepl("^Error in `foo()`:", error_text, fixed = TRUE)

  foo <- function() {
    stop_if_not_in_slide_range(x, 3:4, call = NULL)
  }
  error_text <- tryCatch(foo(), error = paste)
  grepl("^Error:", error_text, fixed = TRUE)
})


test_that("update_named_defaults", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  # errors
  expect_error(
    update_named_defaults(x = 3, default = 1:3),
    "Some default vector elements have no names"
  )
  expect_error(
    update_named_defaults(x = 1:2, default = c(a = 1, b = 2, c = 3)),
    "`x` has incorrect length (2)",
    fixed = TRUE
  )
  expect_error(
    update_named_defaults(x = 1:2, default = c(a = 1)),
    "Length of `x` (2) exceeds length of `default` (1)",
    fixed = TRUE
  )
  expect_error(
    update_named_defaults(
      x = c(d = 3, b = 2, c = 1),
      default = c(aa = 1, bb = 2, cc = 3)
    ),
    'Found 1 unknown name in `x`: "d"',
    fixed = TRUE
  )
  expect_error(
    update_named_defaults(
      x = c(b = 3, b = 2, c = 1),
      default = c(aa = 1, bb = 2, cc = 3)
    ),
    'Duplicate entries in `location`: "bb"',
    fixed = TRUE
  )
  expect_error(
    update_named_defaults(
      x = c(b = 3, b = 2, c = 1),
      default = c(aa = 1, bb = 2, cc = 3),
      argname = "xxx",
      partial = FALSE
    ),
    'Found 3 unknown names in `xxx`: "b", "b", and "c"',
    fixed = TRUE
  )

  # working
  expect_equal(
    update_named_defaults(x = 2, default = c(a = 1, b = 2, c = 3)),
    list(a = 2, b = 2, c = 2)
  )
  expect_equal(
    update_named_defaults(x = 3:1, default = c(a = 1, b = 2, c = 3)),
    list(a = 3, b = 2, c = 1)
  )
  expect_equal(
    update_named_defaults(x = NULL, default = c(a = 1, b = 2, c = 3)),
    list(a = 1, b = 2, c = 3)
  )
  expect_equal(
    update_named_defaults(
      x = NULL,
      default = list(a = 1, b = 2, c = 3),
      as_list = FALSE
    ),
    c(a = 1, b = 2, c = 3)
  )
  expect_equal(
    update_named_defaults(
      x = 3:1,
      default = c(a = 1, b = 2, c = 3),
      as_list = FALSE
    ),
    c(a = 3, b = 2, c = 1)
  )
  expect_equal(
    update_named_defaults(
      x = c(a = 3, b = 2, c = 1),
      default = c(aa = 1, bb = 2, cc = 3)
    ),
    list(aa = 3, bb = 2, cc = 1)
  )
  expect_equal(
    update_named_defaults(x = c(a = 99), default = c(aa = 1, bb = 2, cc = 3)),
    list(aa = 99, bb = 2, cc = 3)
  )
})
