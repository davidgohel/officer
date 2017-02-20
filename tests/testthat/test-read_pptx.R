context("read_pptx")

library(utils)
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
  print(doc, target = "temp.pptx")

  doc <- read_pptx(path = "./temp.pptx")
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
  xml_str <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr/>\n<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Hello world 1</a:t></a:r></a:p></p:txBody></p:sp>"

  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_from_xml(type = "body", value = xml_str)
  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "Hello world 1")
})

unlink("*.pptx")
unlink("*.emf")

