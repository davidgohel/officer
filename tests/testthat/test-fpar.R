get_wml_node <- function(x){
  xml_ <- to_wml(x)
  read_xml( wml_str(xml_) )
}

test_that("fpar", {
  fp <- fpar(
    ftext("toto", shortcuts$fp_bold()),
    ftext("titi", shortcuts$fp_italic())
    )
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
  italic <- xml_find_all(node, "//w:r/w:rPr") %>% xml_child("w:i") %>% xml_attr("val") %in% "true"
  bold <- xml_find_all(node, "//w:r/w:rPr") %>% xml_child("w:b") %>% xml_attr("val") %in% "true"

  expect_equal(italic, c(FALSE, TRUE, FALSE) )
  expect_equal(bold, c(FALSE, FALSE, TRUE) )
})

get_pml_node <- function(x){
  xml_ <- pml_str(to_pml(x))
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
  expect_equal(xml_find_all(node, "//a:r/a:rPr") %>% xml_attr("b"), c("1", "0") )
  expect_equal(xml_find_all(node, "//a:r/a:rPr") %>% xml_attr("i"), c("0", "1") )
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
  node <- read_xml( to_html(fp) )

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
  node <- read_xml( to_html(fp) )
  val <- xml_find_all(node, "//p") %>% xml_attr("style") %>%
    has_css_attr(atname = "text-align", value = "center")
  expect_true(val)
})
