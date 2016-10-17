context("fp_text for wml")

library(utils)
library(xml2)

as_xml_str <- function(str)
paste0( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
        str,
        "</w:document>"
         )

is_italic <- function(doc_){
  !inherits(xml_find_first(doc_, "/w:document/w:rPr/w:i"), "xml_missing")
}
is_bold <- function(doc_){
  !inherits(xml_find_first(doc_, "/w:document/w:rPr/w:b"), "xml_missing")
}
is_underline <- function(doc_){
  !inherits(xml_find_first(doc_, "/w:document/w:rPr/w:u"), "xml_missing")
}

test_that("wml - font size", {
  fp <- fp_text(font.size = 10)
  xml_ <- format(fp, type = "wml")

  doc <- read_xml( as_xml_str(xml_) )

  sz <- xml_find_first(doc, "/w:document/w:rPr/w:sz")
  szCs <- xml_find_first(doc, "/w:document/w:rPr/w:szCs")
  expect_false(inherits( szCs, "xml_missing") )
  expect_false(inherits( sz, "xml_missing") )

  expect_equal(xml_attr(sz, "val"), xml_attr(szCs, "val"))
  expect_equal(xml_attr(sz, "val"), expected = "20")

})

test_that("wml - bold italic underlined", {
  fp_bold <- fp_text(bold = TRUE, italic = FALSE)
  fp_italic <- fp_text(bold = FALSE, italic = TRUE)
  fp_bold_italic <- fp_text(bold = TRUE, italic = TRUE)
  fp_underline <- fp_text(underlined = TRUE )

  xml_bold_ <- format(fp_bold, type = "wml")
  xml_italic_ <- format(fp_italic, type = "wml")
  xml_bolditalic_ <- format(fp_bold_italic, type = "wml")
  xml_underline_ <- format(fp_underline, type = "wml")

  doc_bold_ <- read_xml( as_xml_str(xml_bold_) )
  doc_italic_ <- read_xml( as_xml_str(xml_italic_) )
  doc_bolditalic_ <- read_xml( as_xml_str(xml_bolditalic_) )
  doc_underline_ <- read_xml( as_xml_str(xml_underline_) )

  expect_equal(is_bold(doc_bold_), TRUE)
  expect_equal(is_italic(doc_bold_), FALSE)

  expect_equal(is_bold(doc_italic_), FALSE)
  expect_equal(is_italic(doc_italic_), TRUE)

  expect_equal(is_bold(doc_bolditalic_), TRUE)
  expect_equal(is_italic(doc_bolditalic_), TRUE)

  expect_equal(is_underline(doc_bold_), FALSE)
  expect_equal(is_underline(doc_underline_), TRUE)
})


test_that("wml - font name", {
  fontname = "Arial"
  fp_ <- fp_text(font.family = fontname)

  xml_ <- format(fp_, type = "wml")
  doc_ <- read_xml( as_xml_str(xml_) )

  node <- xml_find_first(doc_, "/w:document/w:rPr/w:rFonts")
  expect_false(inherits(node, "xml_missing"))

  expect_equal( xml_attr(node, "ascii"), fontname )
  expect_equal( xml_attr(node, "hAnsi"), fontname )
  expect_equal( xml_attr(node, "cs"), fontname )
})

test_that("wml - font color", {
  fp_ <- fp_text(color="red")

  xml_ <- format(fp_, type = "wml")
  doc_ <- read_xml( as_xml_str(xml_) )

  node <- xml_find_first(doc_, "/w:document/w:rPr/w:color")
  expect_false(inherits(node, "xml_missing"))

  expect_equal( xml_attr(node, "val"), "FF0000" )
})


test_that("wml - shading color", {
  fp_ <- fp_text(shading.color = rgb(1,0,1))

  xml_ <- format(fp_, type = "wml")
  doc_ <- read_xml( as_xml_str(xml_) )

  node <- xml_find_first(doc_, "/w:document/w:rPr/w:shd")
  expect_false(inherits(node, "xml_missing"))

  expect_equal( xml_attr(node, "fill"), "FF00FF" )
})


