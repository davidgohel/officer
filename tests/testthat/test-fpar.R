context("formatted paragraph")

get_wml_node <- function(x){
  xml_ <- format(x, type = "wml")
  read_xml( wml_str(xml_) )
}

test_that("fpar", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
    )
  expect_output(print(fp), "\\{text:\\{toto\\}\\}\\{text:\\{titi\\}\\}")
  expect_output(print(ftext("toto", shortcuts$fp_bold())), "\\{text:\\{toto\\}\\}")

  fp <- fpar( ftext("toto", fp_text(font.size = 9)), "titi" )
  fp <- update(fp, fp_t = fp_text(bold = TRUE, font.size = 9))
  df <- as.data.frame(fp)
  expect_equal(df$value, c("toto", "titi") )
  expect_equal(df$size, c(9, 9) )
  expect_equal(df$bold, c(FALSE, TRUE) )
  expect_equal(df$italic, c(FALSE, FALSE) )
})


test_that("wml", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
  )
  node <- get_wml_node(fp)

  expect_equal(xml_text(node), "tototiti")
  expect_equal(xml_find_all(node, "//w:r/w:t") %>% xml_text(), c("toto", "titi") )
  expect_equal(xml_find_all(node, "//w:r/w:rPr/*[self::w:i or self::w:b]") %>% xml_name(), c("b", "i") )
  val <- xml_find_all(node, "//w:pPr/w:jc") %>% xml_attr("val")
  expect_equal(val, "left")

  fp <- update(fp, fp_p = fp_par(text.align = "center"))
  node <- get_wml_node(fp)
  val <- xml_find_all(node, "//w:pPr/w:jc") %>% xml_attr("val")
  expect_equal(val, "center")

  fp <- fpar(
    ftext("pi", fp_text()),
    ftext(" = ", fp_text(italic = TRUE)),
    pi
  )
  fp <- update(fp, fp_t = fp_text(bold = TRUE))
  node <- get_wml_node(fp)
  italic <- xml_find_all(node, "//w:r/w:rPr") %>% xml_child("w:i") %>% xml_name() %in% "i"
  bold <- xml_find_all(node, "//w:r/w:rPr") %>% xml_child("w:b") %>% xml_name() %in% "b"

  expect_equal(italic, c(FALSE, TRUE, FALSE) )
  expect_equal(bold, c(FALSE, FALSE, TRUE) )
})

get_pml_node <- function(x){
  xml_ <- pml_str(format(x, type = "pml"))
  read_xml( xml_ )
}

test_that("pml", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
  )
  node <- get_pml_node(fp)

  expect_equal(xml_text(node), "tototiti")
  expect_equal(xml_find_all(node, "//a:r/a:t") %>% xml_text(), c("toto", "titi") )
  expect_equal(xml_find_all(node, "//a:r/a:rPr") %>% xml_attr("b"), c("1", NA) )
  expect_equal(xml_find_all(node, "//a:r/a:rPr") %>% xml_attr("i"), c(NA, "1") )
  val <- xml_find_all(node, "//a:p/a:pPr") %>% xml_attr("algn")
  expect_equal(val, "l")


  fp <- update(fp, fp_p = fp_par(text.align = "center"))
  node <- get_pml_node(fp)
  val <- xml_find_all(node, "//a:p/a:pPr") %>% xml_attr("algn")
  expect_equal(val, "ctr")
})



test_that("css", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
  )
  node <- read_xml( format(fp, type = "html") )

  expect_equal(xml_text(node), "tototiti")
  expect_equal(xml_find_all(node, "//span") %>% xml_text(), c("toto", "titi") )
  bold <- xml_find_all(node, "//span") %>% xml_attr("style") %>%
    has_css_attr(atname = "font-weight", value = "bold")
  expect_equal(bold, c(TRUE, FALSE) )
  italic <- xml_find_all(node, "//span") %>% xml_attr("style") %>%
    has_css_attr(atname = "font-style", value = "italic")
  expect_equal(italic, c(FALSE, TRUE) )

  val <- xml_find_all(node, "//p") %>% xml_attr("style") %>%
    has_css_attr(atname = "text-align", value = "left")
  expect_true(val)


  fp <- update(fp, fp_p = fp_par(text.align = "center"))
  node <- read_xml( format(fp, type = "html") )
  val <- xml_find_all(node, "//p") %>% xml_attr("style") %>%
    has_css_attr(atname = "text-align", value = "center")
  expect_true(val)
})
