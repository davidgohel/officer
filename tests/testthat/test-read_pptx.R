context("read_pptx")

library(xml2)
library(magrittr)

test_that("defaul template", {
  x <- read_pptx()
  expect_equal(length( x ), 0)
  expect_true(file.exists(x$package_dir))
})

test_that("console printing", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  expect_output(print(x), "pptx document with 1 ")
})

test_that("check extention and print document", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  print(x, target = "print.pptx")
  expect_true( file.exists("print.pptx") )

  expect_error(print(x, target = "print.pptxxx"))
})

test_that("check slide_summary", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  sm <- slide_summary(x)
  expect_true(nrow(sm) == 1)
  expect_true(ncol(sm) == 7)
})

test_that("check template", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  print(x, target = "template.pptx")

  expect_silent(x <- read_pptx(path = "template.pptx"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world again")
  print(x, target = "example.pptx")

  expect_equal(length(x), 2)
})



test_that("check slide selection", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 1")
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 2")
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world 3")
  print(x, target = "template.pptx")

  x <- read_pptx(path = "template.pptx")

  x <-  x %>% on_slide(index = 1)
  sm <- slide_summary(x)
  expect_equal(sm$text, "Hello world 1")

  x <-  x %>% on_slide(index = 2)
  sm <- slide_summary(x)
  expect_equal(sm$text, "Hello world 2")

  x <-  x %>% on_slide(index = 3)
  sm <- slide_summary(x)
  expect_equal(sm$text, "Hello world 3")

})

test_that("slide remove", {
  x <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 1") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 2") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 3")
  print(x, target = "template.pptx")

  x <- read_pptx(path = "template.pptx")
  x <- remove_slide(x = x)
  expect_equal(length(x), 2)
  x <- remove_slide(x = x, index = 1)
  expect_equal(length(x), 1)

  sm <- slide_summary(x)
  expect_equal(sm$text, "Hello world 2")
})



test_that("ph remove", {
  x <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 1") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 2") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 3")
  print(x, target = "template.pptx")

  x <- read_pptx(path = "template.pptx")
  x <- ph_remove(x = x)

  expect_equal(nrow(slide_summary(x)), 0)
})

test_that("add text into placeholder", {
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_empty(type = "body")
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "")

  small_red <- fp_text(color = "red", font.size = 14)
  doc <- doc %>%
    ph_add_par(level = 2) %>%
    ph_add_text(str = "chunk 1", style = small_red )
  sm <- slide_summary(doc)
  expect_equal(sm$text, "chunk 1")

  doc <- doc %>%
    ph_add_text(str = "this is ", style = small_red, pos = "before" )
  sm <- slide_summary(doc)
  expect_equal(sm$text, "this is chunk 1")
})

test_that("ph_add_par append when text alrady exists", {
  small_red <- fp_text(color = "red", font.size = 14)
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "This is a ") %>%
    ph_add_par(level = 2) %>%
    ph_add_text(str = "test", style = small_red )
  print(doc, target = "test.pptx")
  sm <- slide_summary(doc)
  expect_equal(sm$text, "This is a test")

  xmldoc <- doc$slide$get_slide(1)$get()
  txt <- xml_find_all(xmldoc, "//p:spTree/p:sp/p:txBody/a:p") %>% xml_text()
  expect_equal(txt, c("This is a ", "test"))
})

test_that("add img into placeholder", {
  img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_img(type = "body", src = img.file,
                       height = 1.06, width = 1.39 )
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "")
  expect_equal(sm$cx, 1.39*914400)
  expect_equal(sm$cy, 1.06*914400)


})

test_that("add xml into placeholder", {
  bold_face <- shortcuts$fp_bold(font.size = 30)
  bold_redface <- update(bold_face, color = "red")

  fpar_ <- fpar(ftext("Hello ", prop = bold_face),
                ftext("World", prop = bold_redface ),
                ftext(", how are you?", prop = bold_face ) )

  doc <- read_pptx() %>%
    add_slide(layout = "Title and Content", master = "Office Theme") %>%
    ph_empty(type = "body") %>%
    ph_add_fpar(value = fpar_, type = "body", level = 2)
  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "Hello World, how are you?")

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  cols <- xml_find_all(xmldoc, "//a:rPr/a:solidFill/a:srgbClr") %>% xml_attr("val")
  expect_equal(cols, c("000000", "FF0000", "000000") )
  expect_equal(xml_find_all(xmldoc, "//a:rPr") %>% xml_attr("b"), rep("1",3))
})


test_that("add xml into placeholder", {
  xml_str <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr/>\n<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Hello world 1</a:t></a:r></a:p></p:txBody></p:sp>"

  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_from_xml(type = "body", value = xml_str)
  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "Hello world 1")
})

test_that("get shape id", {

  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "hello")
  expect_equal(officer:::get_shape_id(doc, type = "body", id_chr = "2"), 1)
  expect_equal(officer:::get_shape_id(doc, type = "body"), 1)
  expect_equal(officer:::get_shape_id(doc, id_chr = "2"), 1)
  expect_error(officer:::get_shape_id(doc, type = "body", id_chr = "1") )
})

#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with_text(x = doc, type = "title", str = "Un titre")
#' slide_summary(doc) # read column id here


unlink("*.pptx")
unlink("*.emf")

