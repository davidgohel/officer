context("fp for cells")

source("utils.R")

test_that("fp_cell", {
  expect_error( fp_cell(text.direction = "sdqsd"), "must be one of" )
  expect_error( fp_cell(margin = -2), "must be a positive integer scalar" )
  expect_error( fp_cell(border = fp_text()), "border must be a fp_border object" )
  expect_error( fp_cell(border.bottom = fp_text()), "border.bottom must be a fp_border object" )
  expect_error( fp_cell(background.color = "#340F2A78d"), " must be a valid color" )
  x <- fp_cell()
  x <- update(x, margin = 3, background.color = "#340F2A78",
              border = fp_border(color = "red"))
  expect_equal(x$background.color, "#340F2A78")
  expect_equal(x$margin.bottom, 3)
  expect_equal(x$margin.left, 3)
  expect_equal(x$border.bottom, fp_border(color = "red"))

  expect_equal(fp_sign( fp_cell() ), "39e7859d6b983b39db6f18cd270bf586" )

})



test_that("print fp_cell", {
  x <- fp_cell()
  expect_output(print(x))
})



pml_cell_node <- function(x){
  xml_ <- pml_str(format(x, type = "pml"))
  doc <- read_xml( xml_ )
  xml_find_first(doc, "//a:tcPr")
}



test_that("pml fp_border", {

  node <- pml_cell_node(fp_cell(background.color = "#00FF0099"))

  col <- xml_child(node, "a:solidFill/a:srgbClr") %>% xml_attr("val")
  expect_equal(col, "00FF00")

  alpha <- xml_child(node, "a:solidFill/a:srgbClr/a:alpha") %>%
    xml_attr("val")
  expect_equal(alpha, "60000")

  border_nodes <- xml_find_all(node, "//*[self::a:lnL or self::a:lnR or self::a:lnT or self::a:lnB]")
  expect_length(border_nodes, 4)

  expect_equal(xml_attr(node, "marB"), "0")
  expect_equal(xml_attr(node, "marT"), "0")
  expect_equal(xml_attr(node, "marR"), "0")
  expect_equal(xml_attr(node, "marL"), "0")

  node <- pml_cell_node(fp_cell(margin = 2, margin.bottom = 4))
  expect_equal(xml_attr(node, "marB") %>% as.integer(), 12700 * 4)
  expect_equal(xml_attr(node, "marT") %>% as.integer(), 12700 * 2)
  expect_equal(xml_attr(node, "marR") %>% as.integer(), 12700 * 2)
  expect_equal(xml_attr(node, "marL") %>% as.integer(), 12700 * 2)

  node <- pml_cell_node(fp_cell(vertical.align = "top"))
  expect_equal(xml_attr(node, "anchor"), "t")
  node <- pml_cell_node(fp_cell(vertical.align = "center"))
  expect_equal(xml_attr(node, "anchor"), "ctr")
  node <- pml_cell_node(fp_cell(vertical.align = "bottom"))
  expect_equal(xml_attr(node, "anchor"), "b")

  x <- fp_cell(text.direction = "btlr")
  node <- pml_cell_node(x)
  expect_equal(xml_attr(node, "vert"), "vert270")
  x <- fp_cell(text.direction = "tbrl")
  node <- pml_cell_node(x)
  expect_equal(xml_attr(node, "vert"), "vert")
})




wml_cell_node <- function(x){
  xml_ <- wml_str(format(x, type = "wml"))
  doc <- read_xml( xml_ )
  xml_find_first(doc, "//w:tcPr")
}

test_that("wml fp_border", {

  node <- wml_cell_node(fp_cell(background.color = "#00FF0099", margin = 2))

  col <- xml_child(node, "w:shd") %>% xml_attr("fill")
  expect_equal(col, "00FF00")

  margins <- xml_child(node, "w:tcMar") %>% xml_children() %>%
    sapply(function(x) xml_attr(x, "w")) %>%
    as.integer()
  expect_equal( margins, rep(40, 4) )

  node <- wml_cell_node(fp_cell(margin = 2, margin.bottom = 0))
  margins <- xml_child(node, "w:tcMar") %>% xml_children() %>%
    sapply(function(x) xml_attr(x, "w")) %>%
    as.integer()
  expect_equal( margins, c(40, 0, 40, 40) )

  node <- wml_cell_node(fp_cell(border = fp_border()))
  border_nodes <- xml_find_all(node, "w:tcBorders/*[self::w:bottom or self::w:top or self::w:left or self::w:right]")
  expect_length(border_nodes, 4)

  node <- wml_cell_node(fp_cell(vertical.align = "top"))
  valign <- xml_child(node, "w:vAlign") %>% xml_attr("val")
  expect_equal(valign, "top")
  node <- wml_cell_node(fp_cell(vertical.align = "center"))
  valign <- xml_child(node, "w:vAlign") %>% xml_attr("val")
  expect_equal(valign, "center")
  node <- wml_cell_node(fp_cell(vertical.align = "bottom"))
  valign <- xml_child(node, "w:vAlign") %>% xml_attr("val")
  expect_equal(valign, "bottom")

  x <- fp_cell(text.direction = "btlr")
  node <- wml_cell_node(x)
  td <- xml_child(node, "w:textDirection") %>% xml_attr("val")
  expect_equal(td, "btLr")
  x <- fp_cell(text.direction = "tbrl")
  node <- wml_cell_node(x)
  td <- xml_child(node, "w:textDirection") %>% xml_attr("val")
  expect_equal(td, "tbRl")
})


test_that("css fp_border", {

  x <- fp_cell(background.color = "#00FF0099", margin = 2)
  expect_true(has_css_color(x, "background-color", "rgba\\(0,255,0,0.60\\)"))

  expect_true(has_css_attr(x, "margin-top", "2pt"))
  expect_true(has_css_attr(x, "margin-bottom", "2pt"))
  expect_true(has_css_attr(x, "margin-left", "2pt"))
  expect_true(has_css_attr(x, "margin-right", "2pt"))

  x <- fp_cell(margin = 2, margin.bottom = 0)
  expect_true(has_css_attr(x, "margin-top", "2pt"))
  expect_true(has_css_attr(x, "margin-bottom", "0pt"))
  expect_true(has_css_attr(x, "margin-left", "2pt"))
  expect_true(has_css_attr(x, "margin-right", "2pt"))

  x <- fp_cell(vertical.align = "top")
  expect_true(has_css_attr(x, "vertical-align", "top"))
  x <- fp_cell(vertical.align = "center")
  expect_true(has_css_attr(x, "vertical-align", "middle"))
  x <- fp_cell(vertical.align = "bottom")
  expect_true(has_css_attr(x, "vertical-align", "bottom"))

})


