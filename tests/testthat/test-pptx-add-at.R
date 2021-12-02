source("utils.R")

test_that("add img into placeholder", {
  skip_on_os("windows")
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- ph_with(doc, value = external_img(img.file), location = ph_location(left = 1, top = 1,
                height = 1.06, width = 1.39) )
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)
  expect_equal(sm$cx, 1.39)
  expect_equal(sm$cy, 1.06)
  expect_equal(sm$offx, 1)
  expect_equal(sm$offy, 1)
})


