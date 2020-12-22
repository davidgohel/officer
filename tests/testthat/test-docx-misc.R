source("utils.R")

test_that("default template", {
  x <- read_docx()
  expect_equal(length( x ), 1)
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


test_that("styles_info is returning a df", {
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

  all_ids <- read_xml(x = file.path(pack_dir, "word/document.xml"))
  all_ids <- xml_find_all(all_ids, "//*[@id]")
  all_ids <- xml_attr(all_ids, "id")

  expect_equal(length(unique(all_ids)), length(all_ids) )
  expect_true( all(grepl("[0-9]+", all_ids )) )

  ids <- as.integer(all_ids)
  expect_true( all( diff(ids) == 1 ) )
})


test_that("cursor behavior", {
  doc <- read_docx()
  doc <- body_add_par(doc, "paragraph 1", style = "Normal")
  doc <- body_add_par(doc, "paragraph 2", style = "Normal")
  doc <- body_add_par(doc, "paragraph 3", style = "Normal")
  doc <- body_bookmark(doc, "bkm1")
  doc <- body_add_par(doc, "paragraph 4", style = "Normal")
  doc <- body_add_par(doc, "paragraph 5", style = "Normal")
  doc <- body_add_par(doc, "paragraph 6", style = "Normal")
  doc <- body_add_par(doc, "paragraph 7", style = "Normal")
  doc <- cursor_begin(doc)
  print(doc, target = "init_doc.docx")

  doc <- read_docx(path = "init_doc.docx")
  doc <- cursor_begin(doc)
  expect_equal( xml_text(doc$doc_obj$get_at_cursor()), "paragraph 1" )
  doc <- cursor_forward(doc)
  expect_equal( xml_text(doc$doc_obj$get_at_cursor()), "paragraph 2" )
  doc <- cursor_end(doc)
  expect_equal( xml_text(doc$doc_obj$get_at_cursor()), "paragraph 7" )
  doc <- cursor_backward(doc)
  expect_equal( xml_text(doc$doc_obj$get_at_cursor()), "paragraph 6" )
  doc <- cursor_reach(doc, keyword = "paragraph 5")
  expect_equal( xml_text(doc$doc_obj$get_at_cursor()), "paragraph 5" )
  doc <- cursor_bookmark(doc, "bkm1")
  expect_equal( xml_text(doc$doc_obj$get_at_cursor()), "paragraph 3" )

})

test_that("cursor and position", {
  doc <- read_docx()
  doc <- body_add_par(doc, "paragraph 1", style = "Normal")
  doc <- body_add_par(doc, "paragraph 2", style = "Normal")
  doc <- cursor_begin(doc)
  doc <- body_add_par(doc, "new 1", style = "Normal", pos = "before")
  doc <- cursor_forward(doc)
  doc <- body_add_par(doc, "new 2", style = "Normal")

  ds_ <- docx_summary(doc)

  expect_equal( ds_$text, c("new 1", "paragraph 1", "new 2", "paragraph 2") )

  doc <- read_docx()
  expect_warning(body_remove(doc), "There is nothing left to remove in the document")

})


unlink("*.docx")
unlink("*.emf")

