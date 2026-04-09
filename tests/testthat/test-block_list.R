fp_bold <- shortcuts$fp_bold()

test_that("block_list structure", {
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  bl <- block_list(
    fpar(ftext("hello world", fp_bold)),
    fpar(
      ftext("hello", fp_bold),
      external_img(src = img.file, height = 1.06, width = 1.39)
    )
  )

  expect_length(bl, 2)
  expect_length(bl[[2]]$chunks, 2)

  expect_is(bl[[2]]$chunks[[1]], "ftext")
  expect_is(bl[[2]]$chunks[[2]], "external_img")
  expect_is(bl[[2]]$chunks[[1]], "cot")
  expect_is(bl[[2]]$chunks[[1]], "run")
  expect_is(bl[[2]]$chunks[[2]], "cot")
})

# block_list_items / list_item ----

test_that("list_item and block_list_items constructors", {
  li <- list_item("hello", level = 2)
  expect_s3_class(li, "list_item")
  expect_s3_class(li$fpar, "fpar")
  expect_equal(li$level, 2L)

  li2 <- list_item(
    fpar(ftext("bold", fp_text(bold = TRUE))),
    level = 1
  )
  expect_s3_class(li2$fpar$chunks[[1]], "ftext")

  items <- block_list_items(
    list_item(fpar("a"), level = 1),
    list_item(fpar("b"), level = 2),
    list_type = "decimal"
  )
  expect_s3_class(items, "block_list_items")
  expect_equal(items$list_type, "decimal")
  expect_length(items$items, 2L)

  expect_error(
    block_list_items("not a list_item"),
    "list_item"
  )
  expect_error(
    block_list_items(list_item("a"), list_type = "invalid"),
    "should be one of"
  )
})

test_that("to_wml.block_list_items produces correct markers and levels", {
  items_bullet <- block_list_items(
    list_item(fpar("Item 1"), level = 1),
    list_item(fpar("Sub-item"), level = 2),
    list_type = "bullet"
  )
  wml <- to_wml(items_bullet)
  expect_true(grepl("officer-list-bullet-", wml))
  expect_true(grepl("w:ilvl w:val=\"0\"", wml))
  expect_true(grepl("w:ilvl w:val=\"1\"", wml))
  expect_false(grepl("w:ind w:left=\"0\"", wml))

  items_decimal <- block_list_items(
    list_item(fpar("First"), level = 1),
    list_item(fpar("Sub-first"), level = 2),
    list_type = "decimal"
  )
  wml2 <- to_wml(items_decimal)
  expect_true(grepl("officer-list-decimal-", wml2))
  expect_true(grepl("w:ilvl w:val=\"1\"", wml2))
})

test_that("to_pml.block_list_items produces correct bullet/number markup", {
  items_bullet <- block_list_items(
    list_item(fpar("Item 1"), level = 1),
    list_item(fpar("Sub-item"), level = 2),
    list_type = "bullet"
  )
  pml <- to_pml(items_bullet)
  expect_true(grepl("a:buChar", pml))
  expect_false(grepl("a:buNone", pml))
  expect_true(grepl("lvl=\"1\"", pml))

  items_decimal <- block_list_items(
    list_item(fpar("First"), level = 1),
    list_item(fpar("Sub-first"), level = 2),
    list_type = "decimal"
  )
  pml2 <- to_pml(items_decimal)
  expect_true(grepl("a:buAutoNum", pml2))
  expect_false(grepl("a:buNone", pml2))
  expect_true(grepl("lvl=\"1\"", pml2))
})

test_that("block_list_items Word round-trip resolves markers", {
  doc <- read_docx()
  bullets <- block_list_items(
    list_item(
      fpar(ftext("Red", fp_text(color = "red"))),
      level = 1
    ),
    list_item(fpar("Sub"), level = 2),
    list_type = "bullet"
  )
  numbers <- block_list_items(
    list_item(fpar("First"), level = 1),
    list_item(fpar("Sub-first"), level = 2),
    list_type = "decimal"
  )
  doc <- body_add(doc, bullets)
  doc <- body_add(doc, numbers)
  out <- print(doc, target = tempfile(fileext = ".docx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)

  doc_xml <- paste(
    readLines(
      file.path(unpack_dir, "word", "document.xml")
    ),
    collapse = ""
  )
  expect_false(grepl("officer-list", doc_xml))

  num_xml <- paste(
    readLines(
      file.path(unpack_dir, "word", "numbering.xml")
    ),
    collapse = ""
  )
  expect_true(grepl("w:val=\"bullet\"", num_xml))
  expect_true(grepl("w:val=\"decimal\"", num_xml))
})

test_that("block_list_items works with write_elements_to_context", {
  doc <- read_docx()
  bullets <- block_list_items(
    list_item(fpar("Bullet 1"), level = 1),
    list_item(fpar("Bullet 2"), level = 2),
    list_type = "bullet"
  )
  numbers <- block_list_items(
    list_item(fpar("Number 1"), level = 1),
    list_item(fpar("Number 2"), level = 1),
    list_type = "decimal"
  )
  ctx <- body_append_start_context(doc)
  ctx <- write_elements_to_context(ctx, bullets, numbers)
  doc <- body_append_stop_context(ctx)
  out <- print(doc, target = tempfile(fileext = ".docx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)

  doc_xml <- paste(
    readLines(
      file.path(unpack_dir, "word", "document.xml")
    ),
    collapse = ""
  )
  expect_false(grepl("officer-list", doc_xml))
})

test_that("body_add_list adds bullet and numbered lists", {
  doc <- read_docx()
  bullets <- block_list_items(
    list_item(fpar("Item 1"), level = 1),
    list_item(fpar("Sub-item"), level = 2),
    list_type = "bullet"
  )
  numbers <- block_list_items(
    list_item(fpar("First"), level = 1),
    list_item(fpar("Second"), level = 1),
    list_type = "decimal"
  )
  doc <- body_add_list(doc, items = bullets)
  doc <- body_add_list(doc, items = numbers)
  out <- print(doc, target = tempfile(fileext = ".docx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)

  doc_xml <- paste(
    readLines(
      file.path(unpack_dir, "word", "document.xml")
    ),
    collapse = ""
  )
  expect_false(grepl("officer-list", doc_xml))
  expect_true(grepl("w:numPr", doc_xml))

  num_xml <- paste(
    readLines(
      file.path(unpack_dir, "word", "numbering.xml")
    ),
    collapse = ""
  )
  expect_true(grepl("w:val=\"bullet\"", num_xml))
  expect_true(grepl("w:val=\"decimal\"", num_xml))
})

test_that("block_list_items PowerPoint output is valid", {
  items <- block_list_items(
    list_item(fpar("Item 1"), level = 1),
    list_item(fpar("Sub-item"), level = 2),
    list_type = "bullet"
  )
  ppt <- read_pptx()
  ppt <- add_slide(ppt, layout = "Title and Content")
  ppt <- ph_with(ppt, items, location = ph_location_type(type = "body"))
  out <- print(ppt, target = tempfile(fileext = ".pptx"))
  expect_true(file.exists(out))
  expect_gt(file.info(out)$size, 0)
})
