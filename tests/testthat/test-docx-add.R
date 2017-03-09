context("add elements in docx")

getncheck <- function(x, str){
  child_ <- xml_child(x, str)
  expect_false( inherits(child_, "xml_missing") )
  child_
}

test_that("body_add_break", {
  x <- read_docx()
  x <- body_add_break(x)

  node <- x$doc_obj$get_at_cursor()
  expect_is( xml_child(node, "/w:r/w:br"), "xml_node" )
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
    body_add_par("paragraph 1") %>%
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

test_that("image add ", {
  img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
  x <- read_docx() %>%
    body_add_par("") %>%
    slip_in_img(src = img.file, style = "strong", width = .3, height = .3)

  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:drawing")
})

test_that("ggplot add", {
  gg_plot <- ggplot(data = iris ) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))
  x <- read_docx() %>%
    body_add_gg(value = gg_plot, style = "centered" )
  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:drawing")
})

test_that("fpar add ", {
  bold_face <- shortcuts$fp_bold(font.size = 20)
  bold_redface <- update(bold_face, color = "red")
  fpar_ <- fpar(ftext("This is a big ", prop = bold_face),
                ftext("text", prop = bold_redface ) ) %>%
    update(fp_p = fp_par(text.align = "center"))
  x <- read_docx() %>% body_add_fpar(fpar_)

  node <- x$doc_obj$get_at_cursor()
  expect_equal(xml_text(node), "This is a big text" )
})


unlink("*.docx")
unlink("*.emf")
