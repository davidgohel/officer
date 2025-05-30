source("utils.R")

test_that("fp_par", {
  expect_error(fp_par(text.align = "glop"), "must be one of")
  expect_error(
    fp_par(padding = -2),
    "padding must be a positive integer scalar"
  )
  expect_error(fp_par(border = fp_text()), "border must be a fp_border object")
  expect_error(
    fp_par(shading.color = "#340F2A78d"),
    "shading\\.color must be a valid color"
  )
  x <- fp_par()
  x <- update(
    x,
    padding = 3,
    text.align = "center",
    shading.color = "#340F2A78",
    border = fp_border(color = "red")
  )
  expect_equal(x$shading.color, "#340F2A78")
  expect_equal(x$padding.bottom, 3)
  expect_equal(x$padding.left, 3)
  expect_equal(x$border.bottom, fp_border(color = "red"))
})


test_that("print fp_par", {
  x <- fp_par()
  expect_output(print(x))
})

get_ppr <- function(x) {
  str <- format(x, type = "pml")
  doc <- as_xml_document(pml_str(str))
  xml_find_first(doc, "a:pPr")
}

test_that("format pml fp_par", {
  x <- fp_par(text.align = "left")
  node <- get_ppr(x)
  expect_equal(xml_attr(node, "algn"), "l")
  x <- update(x, text.align = "center")
  node <- get_ppr(x)
  expect_equal(xml_attr(node, "algn"), "ctr")
  x <- update(x, text.align = "right")
  node <- get_ppr(x)
  expect_equal(xml_attr(node, "algn"), "r")
  x <- update(x, padding = 3, padding.bottom = 0)
  node <- get_ppr(x)
  expect_equal(xml_attr(xml_child(node, "a:spcBef/a:spcPts"), "val"), "300")
  expect_equal(xml_attr(xml_child(node, "a:spcAft/a:spcPts"), "val"), "0")
  expect_equal(xml_attr(node, "marL"), as.character(12700 * 3))
  expect_equal(xml_attr(node, "marR"), as.character(12700 * 3))
})

test_that("format pml fp_par_lite", {
  x <- fp_par_lite()
  expect_equal(ppr_pml(x), "<a:pPr><a:buNone/></a:pPr>")
  x <- fp_par_lite(text.align = "left")
  expect_equal(ppr_pml(x), "<a:pPr algn=\"l\"><a:buNone/></a:pPr>")
})

test_that("format wml fp_par_lite", {
  x <- fp_par_lite()
  expect_equal(ppr_wml(x), "<w:pPr></w:pPr>")
  x <- fp_par_lite(text.align = "left")
  expect_equal(ppr_wml(x), "<w:pPr><w:jc w:val=\"left\"/></w:pPr>")
})


test_that("format css fp_par", {
  x <- fp_par(text.align = "left")
  expect_true(has_css_attr(x, "text-align", "left"))
  x <- update(x, text.align = "center")
  expect_true(has_css_attr(x, "text-align", "center"))
  x <- update(x, text.align = "right")
  expect_true(has_css_attr(x, "text-align", "right"))
  x <- update(x, padding = 3, padding.bottom = 0)
  expect_true(has_css_attr(x, "padding-top", "3pt"))
  expect_true(has_css_attr(x, "padding-bottom", "0pt"))
  expect_true(has_css_attr(x, "padding-right", "3pt"))
  expect_true(has_css_attr(x, "padding-left", "3pt"))

  x <- update(x, shading.color = "#00FF0099")
  expect_true(has_css_color(x, "background-color", "rgba\\(0,255,0,0.60\\)"))
})


is_align <- function(x, align) {
  xml_ <- format(x, type = "wml")
  doc <- read_xml(wml_str(xml_))
  val <- xml_attr(xml_find_first(doc, "/w:document/w:pPr/w:jc"), "val")
  val == align
}


test_that("wml text align", {
  x <- fp_par(text.align = "center")
  expect_true(is_align(x, align = "center"))

  x <- update(x, text.align = "left")
  expect_true(is_align(x, align = "left"))

  x <- update(x, text.align = "right")
  expect_true(is_align(x, align = "right"))

  x <- update(x, text.align = "justify")
  expect_true(is_align(x, align = "both"))
})


get_padding <- function(x, align) {
  xml_ <- format(x, type = "wml")
  doc <- read_xml(wml_str(xml_))
  after <- xml_attr(xml_find_first(doc, "/w:document/w:pPr/w:spacing"), "after")
  before <- xml_attr(
    xml_find_first(doc, "/w:document/w:pPr/w:spacing"),
    "before"
  )
  left <- xml_attr(xml_find_first(doc, "/w:document/w:pPr/w:ind"), "left")
  right <- xml_attr(xml_find_first(doc, "/w:document/w:pPr/w:ind"), "right")
  out <- as.integer(c(after, before, left, right)) / 20
  names(out) <- c("bottom", "top", "left", "right")
  out
}

test_that("wml padding", {
  x <- fp_par(padding = 2)
  expect_equal(
    get_padding(x),
    c("bottom" = 2, "top" = 2, "left" = 2, "right" = 2)
  )

  x <- fp_par(padding = 2, padding.top = 5)
  expect_equal(
    get_padding(x),
    c("bottom" = 2, "top" = 5, "left" = 2, "right" = 2)
  )

  x <- fp_par(padding = 2, padding.bottom = 5)
  expect_equal(
    get_padding(x),
    c("bottom" = 5, "top" = 2, "left" = 2, "right" = 2)
  )
})

test_that("wml shading.color", {
  x <- fp_par(shading.color = "#FF0000")

  xml_ <- format(x, type = "wml")
  doc <- read_xml(wml_str(xml_))
  shd <- xml_attr(xml_find_first(doc, "/w:document/w:pPr/w:shd"), "fill")
  expect_equal(shd, "FF0000")
})

test_that("fp_par_lite", {
  wml <- format(fp_par_lite(), type = "wml")
  expect_equal(wml, "<w:pPr></w:pPr>")
  wml <- format(fp_par_lite(word_style = "dummy"), type = "wml")
  expect_equal(wml, "<w:pPr><w:pStyle w:pstlname=\"dummy\"/></w:pPr>")
  wml <- format(fp_par_lite(text.align = "center"), type = "wml")
  expect_equal(wml, "<w:pPr><w:jc w:val=\"center\"/></w:pPr>")
  wml <- format(fp_par_lite(word_style = "dummy", text.align = "center"), type = "wml")
  expect_equal(wml, "<w:pPr><w:pStyle w:pstlname=\"dummy\"/><w:jc w:val=\"center\"/></w:pPr>")
})



