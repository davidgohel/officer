destfile = tempfile(fileext = ".pptx")
download.file("https://ndownloader.figshare.com/files/16252631",
              destfile = destfile)

test_that("missing tableStyles file in pptx from figshare", {
  testthat::skip_on_os("windows")
  expect_warning({
    x <- read_pptx(destfile)
  })

})

file.remove(destfile)

