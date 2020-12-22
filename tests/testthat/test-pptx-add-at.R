source("utils.R")

test_that("add text into placeholder", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- ph_with(doc, "", location = ph_location(left = 1, top = 1, width = 3, height = 2))
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)

  small_red <- fp_text(color = "red", font.size = 14)
  doc <- ph_add_text(doc, str = "chunk 1", style = small_red, type = "body" )
  sm <- slide_summary(doc)
  expect_equal(sm[1, ]$text, "chunk 1")

  doc <- ph_add_text(doc, str = "this is ", style = small_red, pos = "before", type = "body" )
  sm <- slide_summary(doc)
  expect_equal(sm[1, ]$text, "this is chunk 1")


  xmldoc <- doc$slide$get_slide(1)$get()
  xpath_ <- sprintf("//p:cSld/p:spTree/p:sp")
  node_ <- xml_find_all(xmldoc, xpath_ )
  expect_equal( length(node_), 1)
  node_ <- xml_find_first(xmldoc, "//p:cSld/p:spTree/p:sp" )
  xfrm <- officer:::read_xfrm(node_, file = "qsd", name = "sd" )

  expect_equivalent( unlist(xfrm[c("offx", "offy", "cx", "cy")]) / 914400, c(1,1,3,2))
})


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


