context("fp_text for pml")

library(utils)
library(xml2)

as_xml_str <- function(str)
paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<a:document xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
        str,
        "</a:document>"
         )

is_what <- function(doc_, what = "b"){
  rpr <- xml_find_first(doc_, "/a:document/a:rPr")
  val <- xml_attr(rpr, what)
  !is.na( val ) && val == "1"
}


test_that("pml - font size", {
  fp <- fp_text(font.size = 10)
  xml_ <- format(fp, type = "pml")

  doc_ <- read_xml( as_xml_str(xml_) )
  rpr <- xml_find_first(doc_, "/a:document/a:rPr")

  expect_equal(xml_attr(rpr, "sz"), "1000")
})

test_that("pml - bold italic underlined", {
  fp_bold <- fp_text(bold = TRUE, italic = FALSE)
  fp_italic <- fp_text(bold = FALSE, italic = TRUE)
  fp_bold_italic <- fp_text(bold = TRUE, italic = TRUE)
  fp_underline <- fp_text(underlined = TRUE )

  xml_bold_ <- format(fp_bold, type = "pml")
  xml_italic_ <- format(fp_italic, type = "pml")
  xml_bolditalic_ <- format(fp_bold_italic, type = "pml")
  xml_underline_ <- format(fp_underline, type = "pml")

  doc_bold_ <- read_xml( as_xml_str(xml_bold_) )
  doc_italic_ <- read_xml( as_xml_str(xml_italic_) )
  doc_bolditalic_ <- read_xml( as_xml_str(xml_bolditalic_) )
  doc_underline_ <- read_xml( as_xml_str(xml_underline_) )

  expect_equal(is_what(doc_bold_, "b"), TRUE)
  expect_equal(is_what(doc_bold_, "i"), FALSE)
  expect_equal(is_what(doc_bold_, "u"), FALSE)

  expect_equal(is_what(doc_italic_, "b"), FALSE)
  expect_equal(is_what(doc_italic_, "i"), TRUE)
  expect_equal(is_what(doc_italic_, "u"), FALSE)

  expect_equal(is_what(doc_bolditalic_, "b"), TRUE)
  expect_equal(is_what(doc_bolditalic_, "i"), TRUE)
  expect_equal(is_what(doc_bolditalic_, "u"), FALSE)

  expect_equal(is_what(doc_underline_, "b"), FALSE)
  expect_equal(is_what(doc_underline_, "i"), FALSE)
  expect_equal(is_what(doc_underline_, "u"), TRUE)


})


test_that("pml - font name", {
  fontname = "Arial"
  fp_ <- fp_text(font.family = fontname)

  xml_ <- format(fp_, type = "pml")
  doc_ <- read_xml( as_xml_str(xml_) )

  node <- xml_find_first(doc_, "/a:document/a:rPr/a:latin")
  expect_false(inherits(node, "xml_missing"))
  expect_equal( xml_attr(node, "typeface"), fontname )

  node <- xml_find_first(doc_, "/a:document/a:rPr/a:cs")
  expect_false(inherits(node, "xml_missing"))
  expect_equal( xml_attr(node, "typeface"), fontname )

})

test_that("pml - font color", {
  fp_ <- fp_text(color= rgb(1, 0, 0, .5 ))

  xml_ <- format(fp_, type = "pml")
  doc_ <- read_xml( as_xml_str(xml_) )

  node <- xml_find_first(doc_, "/a:document/a:rPr/a:solidFill/a:srgbClr")
  expect_false(inherits(node, "xml_missing"))
  expect_equal( xml_attr(node, "val"), "FF0000" )

  node <- xml_find_first(doc_, "/a:document/a:rPr/a:solidFill/a:srgbClr/a:alpha")
  expect_equal( xml_attr(node, "val"), "50196" )

})



