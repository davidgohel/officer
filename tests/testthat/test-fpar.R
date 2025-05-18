source("utils.R")

get_wml_node <- function(x) {
  xml_ <- to_wml(x)
  read_xml(wml_str(xml_))
}
get_pml_node <- function(x) {
  xml_ <- pml_str(to_pml(x))
  read_xml(xml_)
}

test_that("fpar content", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
  )
  fp <- fpar(ftext("toto", fp_text(font.size = 9)), "titi")
  fp <- update(fp, fp_t = fp_text(bold = TRUE, font.size = 9))
  df <- as.data.frame(fp)
  expect_equal(df$value, c("toto", "titi"))
  expect_equal(df$size, c(9, 9))
  expect_equal(df$bold, c(FALSE, TRUE))
  expect_equal(df$italic, c(FALSE, FALSE))
})


test_that("fpar wml", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
  )
  node <- get_wml_node(fp)

  expect_equal(xml_text(node), "tototiti")
  expect_equal(xml_text(xml_find_all(node, "//w:r/w:t")), c("toto", "titi"))
  val <- xml_attr(xml_find_all(node, "//w:pPr/w:jc"), "val")
  expect_equal(val, "left")

  fp <- update(fp, fp_p = fp_par(text.align = "center"))
  node <- get_wml_node(fp)
  val <- xml_find_all(node, "//w:pPr/w:jc")
  val <- xml_attr(val, "val")
  expect_equal(val, "center")

  fp <- fpar(
    ftext("pi", fp_text()),
    ftext(" = ", fp_text(italic = TRUE)),
    pi
  )
  fp <- update(fp, fp_t = fp_text(bold = TRUE))
  node <- get_wml_node(fp)
  italic <- xml_find_all(node, "//w:r/w:rPr")
  italic <- xml_child(italic, "w:i")
  italic <- xml_attr(italic, "val") %in% "true"

  bold <- xml_find_all(node, "//w:r/w:rPr")
  bold <- xml_child(bold, "w:b")
  bold <- xml_attr(bold, "val") %in% "true"

  expect_equal(italic, c(FALSE, TRUE, FALSE))
  expect_equal(bold, c(FALSE, FALSE, TRUE))
})


test_that("fpar pml", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
  )
  node <- get_pml_node(fp)

  expect_equal(xml_text(node), "tototiti")
  expect_equal(xml_text(xml_find_all(node, "//a:r/a:t")), c("toto", "titi"))
  expect_equal(xml_attr(xml_find_all(node, "//a:r/a:rPr"), "b"), c("1", "0"))
  expect_equal(xml_attr(xml_find_all(node, "//a:r/a:rPr"), "i"), c("0", "1"))
  val <- xml_attr(xml_find_all(node, "//a:p/a:pPr"), "algn")
  expect_equal(val, "l")

  fp <- update(fp, fp_p = fp_par(text.align = "center"))
  node <- get_pml_node(fp)
  val <- xml_attr(xml_find_all(node, "//a:p/a:pPr"), "algn")
  expect_equal(val, "ctr")
})


test_that("fpar css", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
  )
  node <- read_xml(to_html(fp))

  expect_equal(xml_text(node), "tototiti")
  expect_equal(xml_text(xml_find_all(node, "//span")), c("toto", "titi"))
  bold <- xml_attr(xml_find_all(node, "//span"), "style")
  bold <- has_css_attr(bold, atname = "font-weight", value = "bold")
  expect_equal(bold, c(TRUE, FALSE))

  italic <- xml_attr(xml_find_all(node, "//span"), "style")
  italic <- has_css_attr(italic, atname = "font-style", value = "italic")
  expect_equal(italic, c(FALSE, TRUE))

  val <- xml_attr(xml_find_all(node, "//p"), "style")
  val <- has_css_attr(val, atname = "text-align", value = "left")
  expect_true(val)

  fp <- update(fp, fp_p = fp_par(text.align = "center"))
  node <- read_xml(to_html(fp))
  val <- xml_attr(xml_find_all(node, "//p"), "style")
  val <- has_css_attr(val, atname = "text-align", value = "center")
  expect_true(val)
})
