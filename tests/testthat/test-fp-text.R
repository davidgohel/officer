context("fp for text - format")

source("utils.R")


wml_is_italic <- function(doc_){
  !inherits(xml_find_first(doc_, "/w:document/w:rPr/w:i"), "xml_missing")
}
wml_is_bold <- function(doc_){
  !inherits(xml_find_first(doc_, "/w:document/w:rPr/w:b"), "xml_missing")
}
wml_is_underline <- function(doc_){
  !inherits(xml_find_first(doc_, "/w:document/w:rPr/w:u"), "xml_missing")
}

test_that("wml - font size", {
  fp <- fp_text(font.size = 10)
  xml_ <- format(fp, type = "wml")

  doc <- read_xml( wml_str(xml_) )

  sz <- xml_find_first(doc, "/w:document/w:rPr/w:sz")
  szCs <- xml_find_first(doc, "/w:document/w:rPr/w:szCs")
  expect_false(inherits( szCs, "xml_missing") )
  expect_false(inherits( sz, "xml_missing") )

  expect_equal(xml_attr(sz, "val"), xml_attr(szCs, "val"))
  expect_equal(xml_attr(sz, "val"), expected = "20")

})

test_that("wml - bold italic underlined", {
  fp_bold <- fp_text(bold = TRUE)
  fp_italic <- update(fp_bold, bold = FALSE, italic = TRUE)
  fp_bold_italic <- update(fp_bold, italic = TRUE)
  fp_underline <- fp_text(underlined = TRUE )

  xml_bold_ <- format(fp_bold, type = "wml")
  xml_italic_ <- format(fp_italic, type = "wml")
  xml_bolditalic_ <- format(fp_bold_italic, type = "wml")
  xml_underline_ <- format(fp_underline, type = "wml")

  doc_bold_ <- read_xml( wml_str(xml_bold_) )
  doc_italic_ <- read_xml( wml_str(xml_italic_) )
  doc_bolditalic_ <- read_xml( wml_str(xml_bolditalic_) )
  doc_underline_ <- read_xml( wml_str(xml_underline_) )

  expect_equal(wml_is_bold(doc_bold_), TRUE)
  expect_equal(wml_is_italic(doc_bold_), FALSE)

  expect_equal(wml_is_bold(doc_italic_), FALSE)
  expect_equal(wml_is_italic(doc_italic_), TRUE)

  expect_equal(wml_is_bold(doc_bolditalic_), TRUE)
  expect_equal(wml_is_italic(doc_bolditalic_), TRUE)

  expect_equal(wml_is_underline(doc_bold_), FALSE)
  expect_equal(wml_is_underline(doc_underline_), TRUE)
})


test_that("wml - font name", {
  fontname = "Arial"
  fp_ <- fp_text(font.family = fontname)

  xml_ <- format(fp_, type = "wml")
  doc_ <- read_xml( wml_str(xml_) )

  node <- xml_find_first(doc_, "/w:document/w:rPr/w:rFonts")
  expect_false(inherits(node, "xml_missing"))

  expect_equal( xml_attr(node, "ascii"), fontname )
  expect_equal( xml_attr(node, "hAnsi"), fontname )
  expect_equal( xml_attr(node, "cs"), fontname )
})

test_that("wml - font color", {
  fp_ <- fp_text(color = grDevices::rgb(.8, .2, .1, .6))

  xml_ <- format(fp_, type = "wml")
  doc_ <- read_xml( wml_str(xml_) )

  node <- xml_find_first(doc_, "/w:document/w:rPr/w:color")
  expect_false(inherits(node, "xml_missing"))

  expect_equal( xml_attr(node, "val"), "CC331A" )

  node <- xml_find_first(doc_, "/w:document/w:rPr/w14:textFill")
  expect_false(inherits(node, "xml_missing"))
  expect_equal( xml_attr(xml_child(node, "w14:solidFill/w14:srgbClr"), "val"), "CC331A" )
  expect_equal( xml_attr(xml_child(node, "w14:solidFill/w14:srgbClr/w14:alpha"), "val"), "60000" )

})


test_that("wml - shading color", {
  fp_ <- fp_text(shading.color = rgb(1,0,1))

  xml_ <- format(fp_, type = "wml")
  doc_ <- read_xml( wml_str(xml_) )

  node <- xml_find_first(doc_, "/w:document/w:rPr/w:shd")
  expect_false(inherits(node, "xml_missing"))

  expect_equal( xml_attr(node, "fill"), "FF00FF" )
})



pml_has_true_attr <- function(doc_, what = "b"){
  rpr <- xml_find_first(doc_, "/a:document/a:rPr")
  val <- xml_attr(rpr, what)
  !is.na( val ) && val == "1"
}


test_that("pml - font size", {
  fp <- fp_text(font.size = 10)
  xml_ <- format(fp, type = "pml")

  doc_ <- read_xml( pml_str(xml_) )
  rpr <- xml_find_first(doc_, "/a:document/a:rPr")

  expect_equal(xml_attr(rpr, "sz"), "1000")
})

test_that("pml - bold italic underlined", {
  fp_bold <- fp_text(bold = TRUE)
  fp_italic <- update(fp_bold, bold = FALSE, italic = TRUE)
  fp_bold_italic <- update(fp_bold, italic = TRUE)
  fp_underline <- fp_text(underlined = TRUE )

  xml_bold_ <- format(fp_bold, type = "pml")
  xml_italic_ <- format(fp_italic, type = "pml")
  xml_bolditalic_ <- format(fp_bold_italic, type = "pml")
  xml_underline_ <- format(fp_underline, type = "pml")

  doc_bold_ <- read_xml( pml_str(xml_bold_) )
  doc_italic_ <- read_xml( pml_str(xml_italic_) )
  doc_bolditalic_ <- read_xml( pml_str(xml_bolditalic_) )
  doc_underline_ <- read_xml( pml_str(xml_underline_) )

  expect_equal(pml_has_true_attr(doc_bold_, "b"), TRUE)
  expect_equal(pml_has_true_attr(doc_bold_, "i"), FALSE)
  expect_equal(pml_has_true_attr(doc_bold_, "u"), FALSE)

  expect_equal(pml_has_true_attr(doc_italic_, "b"), FALSE)
  expect_equal(pml_has_true_attr(doc_italic_, "i"), TRUE)
  expect_equal(pml_has_true_attr(doc_italic_, "u"), FALSE)

  expect_equal(pml_has_true_attr(doc_bolditalic_, "b"), TRUE)
  expect_equal(pml_has_true_attr(doc_bolditalic_, "i"), TRUE)
  expect_equal(pml_has_true_attr(doc_bolditalic_, "u"), FALSE)

  expect_equal(pml_has_true_attr(doc_underline_, "b"), FALSE)
  expect_equal(pml_has_true_attr(doc_underline_, "i"), FALSE)
  expect_equal(pml_has_true_attr(doc_underline_, "u"), TRUE)


})


test_that("pml - font name", {
  fontname = "Arial"
  fp_ <- fp_text(font.family = fontname)

  xml_ <- format(fp_, type = "pml")
  doc_ <- read_xml( pml_str(xml_) )

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
  doc_ <- read_xml( pml_str(xml_) )

  node <- xml_find_first(doc_, "/a:document/a:rPr/a:solidFill/a:srgbClr")
  expect_false(inherits(node, "xml_missing"))
  expect_equal( xml_attr(node, "val"), "FF0000" )

  node <- xml_find_first(doc_, "/a:document/a:rPr/a:solidFill/a:srgbClr/a:alpha")
  expect_equal( xml_attr(node, "val"), "50196" )

})






test_that("css", {
  fp <- fp_text(font.size = 10, color = "#00FFFF34", shading.color = "#00FFFFCC")
  expect_true(has_css_color(fp, "color", "rgba\\(0,255,255,0.20\\)"))
  expect_true(has_css_attr(fp, "font-family", "'Arial'"))
  expect_true(has_css_attr(fp, "font-size", "10px"))
  expect_true(has_css_attr(fp, "font-style", "normal"))
  expect_true(has_css_attr(fp, "font-weight", "normal"))
  expect_true(has_css_attr(fp, "text-decoration", "none"))
  expect_true(has_css_color(fp, "background-color", "rgba\\(0,255,255,0.80\\)"))

  fp <- fp_text(bold = TRUE, italic = TRUE, underlined = TRUE)
  expect_true(has_css_attr(fp, "font-style", "italic"))
  expect_true(has_css_attr(fp, "font-weight", "bold"))
  expect_true(has_css_attr(fp, "text-decoration", "underline"))
})

