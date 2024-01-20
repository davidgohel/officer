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

  col <- xml_attr(xml_child(node, "/a:solidFill/a:srgbClr"), "val")
  expect_equal(col, "00FF00")

  alpha <- xml_attr(xml_child(node, "/a:solidFill/a:srgbClr/a:alpha"), "val")
  expect_equal(alpha, "60000")

  type <- xml_attr(xml_child(node, "/a:prstDash"),"val")
  expect_equal(type, "solid")
  node <- pml_border_node(fp_border(style = "dashed"))
  type <- xml_attr(xml_child(node, "/a:prstDash"), "val")
  expect_equal(type, "sysDash")
  node <- pml_border_node(fp_border(style = "dotted", width = 2))
  type <- xml_attr(xml_child(node, "/a:prstDash"), "val")
  expect_equal(type, "sysDot")

  width <- as.integer(xml_attr(node, "w"))
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

  node <- wml_border_node(fp_border(style = "nil"))
  expect_equal(xml_attr(node, "val"), "nil")
  node <- wml_border_node(fp_border(style = "none"))
  expect_equal(xml_attr(node, "val"), "none")
  node <- wml_border_node(fp_border(style = "single"))
  expect_equal(xml_attr(node, "val"), "single")
  node <- wml_border_node(fp_border(style = "thick"))
  expect_equal(xml_attr(node, "val"), "thick")
  node <- wml_border_node(fp_border(style = "double"))
  expect_equal(xml_attr(node, "val"), "double")
  node <- wml_border_node(fp_border(style = "dotDash"))
  expect_equal(xml_attr(node, "val"), "dotDash")
  node <- wml_border_node(fp_border(style = "dotDotDash"))
  expect_equal(xml_attr(node, "val"), "dotDotDash")
  node <- wml_border_node(fp_border(style = "triple"))
  expect_equal(xml_attr(node, "val"), "triple")
  node <- wml_border_node(fp_border(style = "thinThickSmallGap"))
  expect_equal(xml_attr(node, "val"), "thinThickSmallGap")
  node <- wml_border_node(fp_border(style = "thickThinSmallGap"))
  expect_equal(xml_attr(node, "val"), "thickThinSmallGap")
  node <- wml_border_node(fp_border(style = "thinThickThinSmallGap"))
  expect_equal(xml_attr(node, "val"), "thinThickThinSmallGap")
  node <- wml_border_node(fp_border(style = "thinThickMediumGap"))
  expect_equal(xml_attr(node, "val"), "thinThickMediumGap")
  node <- wml_border_node(fp_border(style = "thickThinMediumGap"))
  expect_equal(xml_attr(node, "val"), "thickThinMediumGap")
  node <- wml_border_node(fp_border(style = "thinThickThinMediumGap"))
  expect_equal(xml_attr(node, "val"), "thinThickThinMediumGap")
  node <- wml_border_node(fp_border(style = "thinThickLargeGap"))
  expect_equal(xml_attr(node, "val"), "thinThickLargeGap")
  node <- wml_border_node(fp_border(style = "thickThinLargeGap"))
  expect_equal(xml_attr(node, "val"), "thickThinLargeGap")
  node <- wml_border_node(fp_border(style = "thinThickThinLargeGap"))
  expect_equal(xml_attr(node, "val"), "thinThickThinLargeGap")
  node <- wml_border_node(fp_border(style = "wave"))
  expect_equal(xml_attr(node, "val"), "wave")
  node <- wml_border_node(fp_border(style = "doubleWave"))
  expect_equal(xml_attr(node, "val"), "doubleWave")
  node <- wml_border_node(fp_border(style = "dashSmallGap"))
  expect_equal(xml_attr(node, "val"), "dashSmallGap")
  node <- wml_border_node(fp_border(style = "dashDotStroked"))
  expect_equal(xml_attr(node, "val"), "dashDotStroked")
  node <- wml_border_node(fp_border(style = "threeDEmboss"))
  expect_equal(xml_attr(node, "val"), "threeDEmboss")
  node <- wml_border_node(fp_border(style = "threeDEngrave"))
  expect_equal(xml_attr(node, "val"), "threeDEngrave")
  node <- wml_border_node(fp_border(style = "ridge"))
  expect_equal(xml_attr(node, "val"), "threeDEmboss")
  node <- wml_border_node(fp_border(style = "groove"))
  expect_equal(xml_attr(node, "val"), "threeDEngrave")
  node <- wml_border_node(fp_border(style = "outset"))
  expect_equal(xml_attr(node, "val"), "outset")
  node <- wml_border_node(fp_border(style = "inset"))
  expect_equal(xml_attr(node, "val"), "inset")
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
  x <- fp_border(width = 4, color = col, style = "single")
  rb <- regexp_border(4, "solid", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "dashed")
  rb <- regexp_border(4, "dashed", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "dotted")
  rb <- regexp_border(4, "dotted", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "double")
  rb <- regexp_border(4, "double", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "inset")
  rb <- regexp_border(4, "inset", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "outset")
  rb <- regexp_border(4, "outset", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "none")
  rb <- regexp_border(4, "none", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "nil")
  rb <- regexp_border(4, "none", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "threeDEmboss")
  rb <- regexp_border(4, "ridge", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "ridge")
  rb <- regexp_border(4, "ridge", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "threeDEngrave")
  rb <- regexp_border(4, "groove", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(width = 4, color = col, style = "groove")
  rb <- regexp_border(4, "groove", "rgba\\(0,255,0,0.60\\)", x)
  expect_true( rb )

  x <- fp_border(color = "transparent")
  rb <- regexp_border(1, "solid", "transparent", x)
  expect_true( rb )
})
