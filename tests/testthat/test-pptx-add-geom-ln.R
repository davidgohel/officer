test_that("add outline to placeholder", {
  sp_line_ <- list(
    sp_line(color = "#FFFF00", lwd = 4.5, lty = "solid", linecmpd = "sng"),
    sp_line(color = "black", lwd = 9, lty = "dash", linecmpd = "dbl")
  )
  width <- 4; height <- 3; top <- 7.5 / 2 - height / 2
  left <- c(1, 10 - width - 1);

  loc_ <- lapply(seq_along(sp_line_), function(x) ph_location(left = left[x], top = top, width = width, height = height, bg = "orange", ln = sp_line_[[x]]))

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, "Box with solid outline", location = loc_[[1]])
  doc <- ph_with(doc, "Box with dashed double outline", location = loc_[[2]])

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 2)
  expect_equal(sm$text, c("Box with solid outline", "Box with dashed double outline"))

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  cols <- xml_attr(xml_find_all(xmldoc, "//a:ln/a:solidFill/a:srgbClr"), "val")
  expect_equal(cols, c("FFFF00", "000000") )
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:ln"), "w"), as.character(c(4.5, 9) * 12700))
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:ln"), "cmpd"), c("sng", "dbl"))
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:ln/a:prstDash"), "val"), c("solid", "dash"))
})

test_that("add geometry to placeholder", {
  sp_line_ <- sp_line(color = "#FFFF00", lwd = 4.5, lty = "solid", linecmpd = "sng")

  width <- 4; height <- 3; top <- 7.5 / 2 - height / 2
  left <- c(1, 10 - width - 1);

  loc_ <- list(
    ph_location(left = left[1], top = top, width = width, height = height, bg = "orange", ln = sp_line_, geom = "roundRect"),
    ph_location(left = left[2], top = top, width = width, height = height, bg = "steelblue", ln = sp_line_, geom = "trapezoid")
  )

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, "Round Rectangle", location = loc_[[1]])
  doc <- ph_with(doc, "Trapezoid", location = loc_[[2]])

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 2)
  expect_equal(sm$text, c("Round Rectangle", "Trapezoid"))

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  cols <- xml_attr(xml_find_all(xmldoc, "//p:spPr/a:solidFill/a:srgbClr"), "val")
  expect_equal(cols, c("FFA500", "4682B4") )
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:prstGeom"), "prst"), c("roundRect", "trapezoid"))
  cols <- xml_attr(xml_find_all(xmldoc, "//a:ln/a:solidFill/a:srgbClr"), "val")
  expect_equal(cols, c("FFFF00", "FFFF00") )
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:ln"), "w"), as.character(c(4.5, 4.5) * 12700))
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:ln/a:prstDash"), "val"), c("solid", "solid"))
})

test_that("add line geometry with lineends", {

  headend <- sp_lineend(type = "oval", width = "sm", length = "sm")
  tailend <- sp_lineend(type = "triangle", width = "lg", length = "lg")

  sp_line1 <- sp_line(color = "steelblue", lwd = 4.5, lty = "solid", linecmpd = "sng", headend = headend, tailend = tailend)
  sp_line2 <- update(sp_line1, linecmpd = "dbl", headend = tailend, tailend = headend)

  width <- 4; height <- 3
  top <- 7.5 / 2 - height / 2; left <- c(1, 10 - width - 1)

  loc_ <- list(
    ph_location(left = left[1], top = top, width = width, height = height, ln = sp_line1, geom = "line"),
    ph_location(left = left[2], top = top, width = width, height = height, ln = sp_line2, rotation = 180, geom = "line")
  )

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, "", location = loc_[[1]])
  doc <- ph_with(doc, "", location = loc_[[2]])

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 2)

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:prstGeom"), "prst"), c("line", "line"))
  cols <- xml_attr(xml_find_all(xmldoc, "//a:ln/a:solidFill/a:srgbClr"), "val")
  expect_equal(cols, c("4682B4", "4682B4") )
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:ln"), "cmpd"), c("sng", "dbl"))
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:headEnd"), "type"), c("oval", "triangle"))
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:tailEnd"), "type"), c("triangle", "oval"))
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:headEnd"), "w"), c("sm", "lg"))
  expect_equal( xml_attr(xml_find_all(xmldoc, "//a:tailEnd"), "len"), c("lg", "sm"))
})
