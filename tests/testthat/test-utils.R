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
  df <- df_rename(mtcars, c("mpg", "cyl"), c("A", "B"))
  expect_true(all(names(df)[1:2] == c("A", "B")))
})
