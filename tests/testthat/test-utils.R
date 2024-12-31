test_that("trailing file index extraction / sorting", {
  files <- c("slideLayout1.xml", "slideLayout11.xml", "slideLayout2.xml", "slideLayout10.xml")

  expect_equal(get_file_index(files), c(1, 11, 2, 10))

  expect_equal(sort_vec_by_index(files), c("slideLayout1.xml", "slideLayout2.xml", "slideLayout10.xml", "slideLayout11.xml"))

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
