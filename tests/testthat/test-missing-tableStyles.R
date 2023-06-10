test_that("missing tableStyles file in pptx from figshare", {
  testthat::skip_if_offline()
  testthat::skip_on_os("windows")

  destfile <- tempfile(fileext = ".pptx")
  on.exit(file.remove(destfile), add = TRUE)

  download.file(
    "https://ndownloader.figshare.com/files/16252631",
    destfile = destfile
  )

  expect_warning({
    x <- read_pptx(destfile)
  })
})
