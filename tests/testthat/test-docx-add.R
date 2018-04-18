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


test_that("body_end_sections", {

  x <- read_docx() %>%
    body_add_par("paragraph 1", style = "Normal") %>%
    body_end_section_landscape()

  node <- x$doc_obj$get_at_cursor()
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )
  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false( inherits(ps, "xml_missing") )
  expect_equal( xml_attr(ps, "orient"), "landscape")

  x <- body_add_par(x, "paragraph 1", style = "Normal") %>%
    body_add_par("paragraph 2", style = "Normal") %>%
    body_end_section_columns()

  node <- x$doc_obj$get_at_cursor()
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )
  sect <- xml_child(node, "w:pPr/w:sectPr")

  expect_false( inherits(sect, "xml_missing") )
  expect_equal( xml_attr(xml_child(sect, "w:type"), "val"), "continuous")
  expect_false( inherits(xml_child(sect, "w:cols"), "xml_missing") )

  x <- body_add_par(x, "paragraph 1", style = "Normal") %>%
    body_add_par("paragraph 2", style = "Normal") %>%
    body_end_section_columns_landscape()

  node <- x$doc_obj$get_at_cursor()
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )

  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false( inherits(ps, "xml_missing") )
  expect_equal( xml_attr(ps, "orient"), "landscape")

  sect <- xml_child(node, "w:pPr/w:sectPr")
  expect_false( inherits(sect, "xml_missing") )
  expect_false( inherits(xml_child(sect, "w:cols"), "xml_missing") )

  x <- body_add_par(x, "paragraph 1", style = "Normal") %>%
    body_add_par("paragraph 2", style = "Normal") %>%
    body_end_section_portrait()

  node <- x$doc_obj$get_at_cursor()
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )

  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false( inherits(ps, "xml_missing") )
  expect_equal( xml_attr(ps, "orient"), "portrait")

  xml_flags <- x$doc_obj$get() %>%
    xml_find_all("//w:officersection")
  expect_length(xml_flags, 4)
  docx_file <- tempfile(fileext = ".docx")
  docx_dir <- tempfile()
  print(x, target = docx_file)
  unpack_folder(docx_file, docx_dir)



  xml_flags <- read_xml(file.path(docx_dir, "word/document.xml")) %>%
    xml_find_all("//w:officersection")
  expect_length(xml_flags, 0)

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

test_that("body_add_img", {

  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  x <- read_docx() %>%
    body_add_img(img.file, width=2.5, height=1.3)

  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:drawing")
})

test_that("slip_in_img", {
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  x <- read_docx() %>%
    body_add_par("") %>%
    slip_in_img(src = img.file, style = "strong", width = .3, height = .3)

  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:drawing")
})

test_that("ggplot add", {
  testthat::skip_if_not(requireNamespace("ggplot2", quietly = TRUE))
  library("ggplot2")

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

test_that("add docx into docx", {

  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  doc <- read_docx()
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
  print(doc, target = "external_file.docx")

  final_doc <- read_docx()
  doc <- body_add_docx(x = doc, src = "external_file.docx" )
  print(doc, target = "final.docx")

  new_dir <- tempfile()
  unpack_folder("final.docx", folder = new_dir)

  doc_parts <- read_xml(file.path(new_dir, "[Content_Types].xml")) %>%
    xml_find_all("d1:Override[@ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']") %>%
    xml_attr("PartName") %>% basename()
  expect_equal(doc_parts[grepl("\\.docx$", doc_parts)], list.files(new_dir, pattern = "\\.docx$") )
})


unlink("*.docx")
unlink("*.emf")
