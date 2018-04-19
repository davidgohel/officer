context("fp for borders")

source("utils.R")

test_that("fp_border", {
  expect_error( fp_border(width = -5), "width must be a positive numeric scalar" )
  expect_error( fp_border(color = "glop"), "color must be a valid color" )
  expect_error( fp_border(style = "glop"), "style must be one of" )
  x <- fp_border(color = "red", style = "dashed", width = 5)
  x <- update(x, color = "yellow", style = "solid", width = 2)
  expect_equal(x$color, "yellow")
  expect_equal(x$style, "solid")
  expect_equal(x$width, 2)


})

pml_border_node <- function(x, side = "L"){
  x <- fp_cell(border = x )
  xml_ <- pml_str(format(x, type = "pml"))
  doc <- read_xml( xml_ )
  xml_find_first(doc, paste0("//a:tcPr/a:ln", side))
}



test_that("pml fp_border", {

  node <- pml_border_node(fp_border(width = 2, color = "#00FF0099"))

  col <- xml_child(node, "/a:solidFill/a:srgbClr") %>% xml_attr("val")
  expect_equal(col, "00FF00")

  alpha <- xml_child(node, "/a:solidFill/a:srgbClr/a:alpha") %>%
    xml_attr("val")
  expect_equal(alpha, "60000")

  type <- xml_child(node, "/a:prstDash") %>%
    xml_attr("val")
  expect_equal(type, "solid")
  node <- pml_border_node(fp_border(style = "dashed"))
  type <- xml_child(node, "/a:prstDash") %>%
    xml_attr("val")
  expect_equal(type, "sysDash")
  node <- pml_border_node(fp_border(style = "dotted", width = 2))
  type <- xml_child(node, "/a:prstDash") %>%
    xml_attr("val")
  expect_equal(type, "sysDot")

  width <- xml_attr(node, "w") %>%
    as.integer()
  expect_equal(width, 12700 * 2)

  node <- pml_border_node(fp_border(color = "transparent"))
  col <- xml_child(node, "a:noFill")
  expect_false(inherits(col, "xml_missing"))
})


wml_border_node <- function(x, side = "bottom"){
  x <- fp_cell(border = x )
  xml_ <- wml_str(format(x, type = "wml"))
  doc <- read_xml( xml_ )
  xml_find_first(doc, paste0("//w:tcPr/w:tcBorders/w:", side))
}

test_that("wml fp_border", {

  node <- wml_border_node(fp_border(width = 2, color = "#00FF00"))

  col <- xml_attr(node, "color")
  expect_equal(col, "00FF00")

  sz <- xml_attr(node, "sz")
  expect_equal(sz, "16")

  node <- wml_border_node(fp_border(style = "dotted"))
  expect_equal(xml_attr(node, "val"), "dotted")
  node <- wml_border_node(fp_border(style = "solid"))
  expect_equal(xml_attr(node, "val"), "single")
  node <- wml_border_node(fp_border(style = "dashed"))
  expect_equal(xml_attr(node, "val"), "dashed")

})


regexp_border <- function(width, style, color, x){
  css <- format(fp_par(border = x), type = "html")
  reg <- sprintf("border-bottom: %.02fpt %s %s", width, style, color)
  grepl(reg, css)
}

test_that("css fp_border", {

  col <- "#00FF00"
  x <- fp_border(width = 2, color = col, style = "solid")
  rb <- regexp_border(2, "solid", "rgba\\(0,255,0,1.00\\)", x)
  expect_true( rb )

  col <- "#00FF0099"
  x <- fp_border(width = 4, color = col, style = "dashed")
  rb <- regexp_border(4, "dashed", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "dotted")
  rb <- regexp_border(4, "dotted", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(color = "transparent")
  rb <- regexp_border(1, "solid", "transparent", x)
  expect_true( rb )
})
