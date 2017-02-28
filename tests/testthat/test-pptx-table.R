context("table in pptx")

test_that("ph_with_table", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", master = "Office Theme")
  x <- ph_with_table(x, value = iris, type = "body")
  xmldoc <- x$slide$get_slide(1)$get()
  xml_tab <- xml_find_all(xmldoc, xpath = "p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl")
  expect_false( inherits(xml_tab, "xml_missing") )
})

test_that("pml structure", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", master = "Office Theme")
  x <- ph_with_table(x, value = mtcars, type = "body")
  xmldoc <- x$slide$get_slide(1)$get()

  xml_tr <- xml_find_all(xmldoc, xpath = "p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr")
  expect_length(xml_tr, nrow(mtcars) + 1)
  expect_equal(xml_attr(xml_tr, "h") %>% unique() %>% length(), 1)
  expect_true( all( grepl( "^[0-9]+$", xml_attr(xml_tr, "h") ) ) )

  xml_tr <- xml_find_all(xmldoc, xpath = "p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc")
  expect_length(xml_tr, (nrow(mtcars) + 1) * ncol(mtcars))

})

