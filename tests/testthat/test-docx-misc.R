context("miscellaneous check for docx")

source("utils.R")

test_that("default template", {
  x <- read_docx()
  expect_equal(length( x ), 2)
  expect_true(file.exists(x$package_dir))
})

test_that("docx dim", {
  x <- read_docx()
  dims <- docx_dim(x)

  expect_equal( names(dims), c("page", "landscape", "margins") )
  expect_length( dims$page, 2 )
  expect_length( dims$landscape, 1 )
  expect_length( dims$margins, 6 )

  expect_equal( dims$page, c(width=8.263889, height=11.694444), tolerance = .001 )

})

test_that("list bookmarks", {
  template_file <- system.file(package = "officer", "doc_examples/example.docx")
  x <- read_docx(path = template_file)
  bookmarks <- docx_bookmarks(x)

  expect_equal( bookmarks, c("bmk_1", "bmk_2") )
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

test_that("style is read from document", {
  x <- read_docx()
  expect_silent({
    x <- body_add_par(x = x, value = "paragraph 1", style = "Normal")
  })

  expect_error({
    x <- body_add_par(x = x, value = "paragraph 1", style = "blahblah")
  })
})


test_that("styles_info is returning a tidy df", {
  x <- read_docx()
  df <- styles_info(x)

  expect_is( df, "data.frame" )
  expect_true( all( c("style_type", "style_id", "style_name", "is_custom", "is_default") %in% names(df)) )
  expect_is( df$style_type, "character" )
  expect_is( df$style_id, "character" )
  expect_is( df$style_name, "character" )
  expect_is( df$is_custom, "logical" )
  expect_is( df$is_default, "logical" )

})


test_that("id are sequentially defined", {
  doc <- read_docx()
  any_img <- FALSE
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  if( file.exists(img.file) ){
    doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
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


test_that("cursor behavior", {
  doc <- read_docx() %>%
    body_add_par("paragraph 1", style = "Normal") %>%
    body_add_par("paragraph 2", style = "Normal") %>%
    body_add_par("paragraph 3", style = "Normal") %>%
    body_bookmark("bkm1") %>%
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
  doc <- doc %>% cursor_reach(keyword = "paragraph 5")
  expect_equal( doc$doc_obj$get_at_cursor() %>% xml_text(), "paragraph 5" )
  doc <- doc %>% cursor_bookmark("bkm1")
  expect_equal( doc$doc_obj$get_at_cursor() %>% xml_text(), "paragraph 3" )

})

test_that("body remove", {
  doc <- read_docx() %>%
    body_add_par("paragraph 1", style = "Normal") %>%
    body_add_par("paragraph 2", style = "Normal") %>%

    cursor_begin() %>%
    body_remove() %>%

    body_add_par("new 1", style = "Normal") %>%
    cursor_forward() %>%
    body_add_par("new 2", style = "Normal")

  ds_ <- docx_summary(doc)

  expect_equal( ds_$text, c("paragraph 1", "new 1", "paragraph 2", "new 2") )

  doc <- read_docx()
  doc <- body_remove(doc)
  expect_warning(body_remove(doc), "There is nothing left to remove in the document")

})


unlink("*.docx")
unlink("*.emf")

