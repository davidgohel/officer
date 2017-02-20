context("paragraph formatting properties")

library(xml2)
library(magrittr)

test_that("fp_par", {
  expect_error( fp_par(text.align = "glop"), "must be one of" )
  expect_error( fp_par(padding = -2), "padding must be a positive integer scalar" )
  expect_error( fp_par(border = fp_text()), "border must be a fp_border object" )
  expect_error( fp_par(shading.color = "#340F2A78d"), "shading\\.color must be a valid color" )
  x <- fp_par()
  x <- update(x, padding = 3,
              text.align = "center",
              shading.color = "#340F2A78", border = fp_border(color = "red"))
  expect_equal(x$shading.color, "#340F2A78")
  expect_equal(x$padding.bottom, 3)
  expect_equal(x$padding.left, 3)
  expect_equal(x$border.bottom, fp_border(color = "red"))

  x <- fp_par(padding = 0, border = shortcuts$b_null())
  expect_identical(dim(x), c("width" = 0, "height" = 0) )
})



test_that("print fp_par", {
  x <- fp_par()
  expect_output(print(x))
})

get_ppr <- function(x){
  str <- format(x, type = "pml")
  str <- paste0("<a:p ", officer:::pml_ns, ">",
                str, "</a:p>")
  doc <- as_xml_document(str)
  xml_find_first(doc, "a:pPr")
}

test_that("format pml fp_par", {
  x <- fp_par(text.align = "left")
  node <- get_ppr(x)
  expect_equal( xml_attr(node, "algn"), "l" )
  x <- update(x, text.align = "center")
  node <- get_ppr(x)
  expect_equal( xml_attr(node, "algn"), "ctr" )
  x <- update(x, text.align = "right")
  node <- get_ppr(x)
  expect_equal( xml_attr(node, "algn"), "r" )
  x <- update(x, padding = 3, padding.bottom = 0)
  node <- get_ppr(x)
  expect_equal( xml_child(node, "a:spcBef/a:spcPts") %>% xml_attr("val"), "300" )
  expect_equal( xml_child(node, "a:spcAft/a:spcPts") %>% xml_attr("val"), "0" )
  expect_equal( xml_attr(node, "marL"), as.character(12700*3) )
  expect_equal( xml_attr(node, "marR"), as.character(12700*3) )
})




unlink("*.pptx")

