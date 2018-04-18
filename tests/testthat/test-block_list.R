context("block list")

fp_bold <- shortcuts$fp_bold()

test_that("block_list structure", {

  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  bl <- block_list(
    fpar(ftext("hello world", fp_bold)),
    fpar( ftext("hello", fp_bold),
          stext(" world", "strong"),
          external_img(src = img.file, height = 1.06, width = 1.39)
          )
    )

  expect_length(bl, 2)
  expect_length(bl[[2]]$chunks, 3)

  expect_is(bl[[2]]$chunks[[1]], "ftext")
  expect_is(bl[[2]]$chunks[[2]], "stext")
  expect_is(bl[[2]]$chunks[[3]], "external_img")
  expect_is(bl[[2]]$chunks[[1]], "cot")
  expect_is(bl[[2]]$chunks[[2]], "cot")
  expect_is(bl[[2]]$chunks[[3]], "cot")
})

test_that("print block_list", {
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  bl <- block_list(
    fpar( ftext("hello", fp_bold),
          stext(" world", "strong"),
          external_img(src = img.file, height = 1.06, width = 1.39)
          )
    )
  expect_output(print(bl), "{text:{hello}}{text:{ world}}{image:", fixed = TRUE)
})

