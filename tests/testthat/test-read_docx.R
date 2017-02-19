context("read_docx")

library(utils)
library(xml2)
library(magrittr)

test_that("defaul template", {
  x <- read_docx()
  expect_equal(length( x ), 2)
  expect_true(file.exists(x$doc_obj$package_dirname()))
})

test_that("console printing", {
  x <- read_docx()
  x <- body_add_par(x, "Hello world", style = "Normal")
  expect_output(print(x), "docx document with")
})

test_that("check extention and print document", {
  x <- read_docx()
  print(x, target = "print.docx")
  expect_true( file.exists("print.docx") )

  expect_error(print(x, target = "print.docxxx"))
})


test_that("id are sequentially defined", {

  doc <- read_docx()
  any_img <- FALSE
  img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
  if( file.exists(img.file) ){
    doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
    any_img <- TRUE
  }
  if( require("ionicons") ){
    calendar_src = as_png(name = "calendar", fill = "#FFE64D", width = 144, height = 144)
    doc <- body_add_img(x = doc, src = calendar_src, height = 2, width = 2 )
    any_img <- TRUE
  }
  if( require("devEMF") ){
    emf("bar.emf", height = 5, width = 5)
    barplot(1:10, col = 1:10)
    dev.off()
    doc <- body_add_img(x = doc, src = "bar.emf", height = 5, width = 5)
    any_img <- TRUE
  }
  print(doc, target = "body_add_img.docx" )
  skip_if_not(any_img)

  pack_dir <- tempfile(pattern = "dir")
  unpack_folder(file = "body_add_img.docx", folder = pack_dir)

  all_ids <- read_xml(x = file.path(pack_dir, "word/document.xml")) %>%
    xml_find_all("//*[@id]") %>% xml_attr("id")

  expect_equal(length(unique(all_ids)), length(all_ids) )
  expect_true( all(grepl("[0-9]+", all_ids )) )

  ids <- as.integer(all_ids)
  expect_true( all( diff(ids) == 1 ) )
})


test_that("id are sequentially defined", {

  doc <- read_docx() %>%
    body_add_par("paragraph 1", style = "Normal") %>%
    body_add_par("paragraph 2", style = "Normal") %>%
    body_add_par("paragraph 3", style = "Normal") %>%
    body_add_par("paragraph 4", style = "Normal") %>%
    body_add_par("paragraph 5", style = "Normal") %>%
    body_add_par("paragraph 6", style = "Normal") %>%
    body_add_par("paragraph 7", style = "Normal") %>%
    cursor_begin() %>% body_remove() %>%
    print(target = "init_doc.docx")

  doc <- read_docx(path = "init_doc.docx") %>%
    cursor_begin()
  expect_equal( doc$doc_obj$get_at_cursor() %>% xml_text(), "paragraph 1" )
  doc <- doc %>% cursor_forward()
  expect_equal( doc$doc_obj$get_at_cursor() %>% xml_text(), "paragraph 2" )
  doc <- doc %>% cursor_end()
  expect_equal( doc$doc_obj$get_at_cursor() %>% xml_text(), "paragraph 7" )
  doc <- doc %>% cursor_backward()
  expect_equal( doc$doc_obj$get_at_cursor() %>% xml_text(), "paragraph 6" )

})

unlink("*.docx")
unlink("*.emf")
