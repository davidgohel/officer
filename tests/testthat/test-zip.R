context("folder pack")


test_that("pack_folder does not use full path", {
  dir_ <- tempfile()
  dir.create(dir_)
  file <- tempfile(tmpdir = dir_)
  cat("test", file = file)
  pack_folder(dir_, target = "test.zip")

  dir_ <- tempfile()
  unpack_folder(file = "test.zip", dir_ )
  expect_equal(list.files(dir_), basename(file))
  unlink("test.zip", force = TRUE)
})

test_that("pack_folder behavior", {
  dir_ <- tempfile()
  dir.create(dir_)
  file <- tempfile(tmpdir = dir_)
  cat("test", file = file)

  expect_error(pack_folder(dir_, target = "dummy_dir/test.zip"))
})

