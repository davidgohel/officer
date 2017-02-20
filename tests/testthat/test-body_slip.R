context("body_add")

library(utils)
library(xml2)
library(magrittr)

getncheck <- function(x, str){
  child_ <- xml_child(x, str)
  expect_false( inherits(child_, "xml_missing") )
  child_
}


test_that("style is read from document", {
  x <- read_docx()
  expect_silent({
    x <- body_add_par(x = x, value = "paragraph 1", style = "Normal")
  })

  expect_error({
    x <- body_add_par(x = x, value = "paragraph 1", style = "blahblah")
  })
})

test_that("body_add_break", {
  x <- read_docx()
  x <- body_add_break(x)

  node <- x$doc_obj$get_at_cursor()
  expect_is( xml_child(node, "/w:r/w:br"), "xml_node" )
})

test_that("body_add_table", {
  x <- read_docx()
  x <- body_add_table(x, value = iris, style = "table_template")
  node <- x$doc_obj$get_at_cursor()
  expect_equal( xml_name(node), "tbl" )
})

test_that("body_add_section", {

  x <- read_docx() %>%
    body_add_par("paragraph 1", style = "Normal") %>%
    body_end_section(landscape = TRUE)

  node <- x$doc_obj$get_at_cursor()
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )
  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false( inherits(ps, "xml_missing") )
  expect_equal( xml_attr(ps, "orient"), "landscape")
})


test_that("body_add_toc", {

  x <- read_docx() %>%
    body_add_par("paragraph 1", style = "Normal") %>%
    body_add_toc()

  node <- x$doc_obj$get_at_cursor()

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")
  expect_equal( xml_attr(child_, "dirty"), "true")

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")
  expect_equal( xml_attr(child_, "dirty"), "true")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal( xml_attr(child_, "dirty"), "true")
  expect_equal( xml_text(child_), "TOC \\o \"1-3\" \\h \\z \\u" )


  x <- x %>%
    body_add_toc(style = "Normal")
  node <- x$doc_obj$get_at_cursor()

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")
  expect_equal( xml_attr(child_, "dirty"), "true")
  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")
  expect_equal( xml_attr(child_, "dirty"), "true")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal( xml_attr(child_, "dirty"), "true")
  expect_equal( xml_text(child_), "TOC \\h \\z \\t \"Normal;1\"" )

})

test_that("seqfield add ", {
  x <- read_docx() %>%
    body_add_par("Time is: ", style = "Normal") %>%
    slip_in_seqfield(
      str = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT",
      style = 'strong')

  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")
  getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal( xml_attr(child_, "dirty"), "true")
  expect_equal( xml_text(child_), "TIME \\@ \"HH:mm:ss\" \\* MERGEFORMAT" )

  x <- body_add_par(x, " - This is a figure title", style = "centered") %>%
    slip_in_seqfield(str = "SEQ Figure \u005C* roman",
                     style = 'Default Paragraph Font', pos = "before") %>%
    slip_in_text("Figure: ", style = "strong", pos = "before")
  node <- x$doc_obj$get_at_cursor()
  expect_equal( xml_text(node), "Figure: SEQ Figure \\* roman - This is a figure title" )
})


test_that("image add ", {
  img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
  x <- read_docx() %>%
    body_add_par("", style = "Normal") %>%
    slip_in_img(src = img.file, style = "strong", width = .3, height = .3)

  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:drawing")
})


unlink("*.docx")
unlink("*.emf")
