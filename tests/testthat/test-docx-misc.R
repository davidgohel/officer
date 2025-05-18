test_that("default template", {
  x <- read_docx()
  expect_equal(length(x), 1)
  expect_true(file.exists(x$package_dir))
})

test_that("docx dim", {
  x <- read_docx()
  dims <- docx_dim(x)

  expect_equal(names(dims), c("page", "landscape", "margins"))
  expect_length(dims$page, 2)
  expect_length(dims$landscape, 1)
  expect_length(dims$margins, 6)

  expect_equal(
    dims$page,
    c(width = 8.263889, height = 11.694444),
    tolerance = .001
  )

  ps <- prop_section(
    page_size = page_size(
      width = 8.263889,
      height = 11.694444,
      orient = "landscape"
    ),
    page_margins = page_mar(top = 2),
    type = "oddPage"
  )
  x <- read_docx()
  x <- body_add_par(x = x, value = "paragraph 1", style = "Normal")
  x <- body_end_block_section(x, block_section(ps))
  x <- cursor_begin(x)
  dims <- docx_dim(x)
  expect_equivalent(
    object = dims,
    expected = list(
      page = c(width = 11.694444, height = 8.263889),
      landscape = TRUE,
      margins = c(
        top = 2,
        bottom = 1417 / 1440,
        left = 1417 / 1440,
        right = 1417 / 1440,
        header = 708 / 1440,
        footer = 708 / 1440
      )
    ),
    tolerance = .00001
  )
})

test_that("list bookmarks", {
  template_file <- system.file(package = "officer", "doc_examples/example.docx")
  x <- read_docx(path = template_file)
  bookmarks <- docx_bookmarks(x)

  expect_equal(bookmarks, c("bmk_1", "bmk_2"))
})

test_that("console printing", {
  x <- read_docx()
  x <- body_add_par(x, "Hello world", style = "Normal")
  expect_output(print(x), "docx document with")
})

test_that("check extention and print document", {
  x <- read_docx()
  outfile <- print(x, target = tempfile(fileext = ".docx"))
  expect_true(file.exists(outfile))

  expect_error(print(x, target = tempfile(fileext = ".docxxxx")))
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


test_that("id are sequentially defined", {
  doc <- read_docx()
  any_img <- FALSE
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  if (file.exists(img.file)) {
    doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39)
    any_img <- TRUE
  }
  tmp_file <- print(doc, target = tempfile(fileext = ".docx"))
  skip_if_not(any_img)

  pack_dir <- tempfile(pattern = "dir")
  unpack_folder(file = tmp_file, folder = pack_dir)

  all_ids <- read_xml(x = file.path(pack_dir, "word/document.xml"))
  all_ids <- xml_find_all(all_ids, "//*[@id]")
  all_ids <- xml_attr(all_ids, "id")

  expect_equal(length(unique(all_ids)), length(all_ids))
  expect_true(all(grepl("[0-9]+", all_ids)))

  ids <- as.integer(all_ids)
  expect_true(all(diff(ids) == 1))
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
  init_file <- print(doc, target = tempfile(fileext = ".docx"))

  doc <- read_docx(path = init_file)
  doc <- cursor_begin(doc)
  current_node <- docx_current_block_xml(doc)
  expect_equal(xml_text(current_node), "paragraph 1")
  doc <- cursor_forward(doc)
  current_node <- docx_current_block_xml(doc)
  expect_equal(xml_text(current_node), "paragraph 2")
  doc <- cursor_end(doc)
  current_node <- docx_current_block_xml(doc)
  expect_equal(xml_text(current_node), "paragraph 7")
  doc <- cursor_backward(doc)
  current_node <- docx_current_block_xml(doc)
  expect_equal(xml_text(current_node), "paragraph 6")
  doc <- cursor_reach(doc, keyword = "paragraph 5")
  current_node <- docx_current_block_xml(doc)
  expect_equal(xml_text(current_node), "paragraph 5")
  doc <- cursor_bookmark(doc, "bkm1")
  current_node <- docx_current_block_xml(doc)
  expect_equal(xml_text(current_node), "paragraph 3")
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

  expect_equal(ds_$text, c("new 1", "paragraph 1", "new 2", "paragraph 2"))
  doc <- body_remove(doc)
  doc <- body_remove(doc)
  ds_ <- docx_summary(doc)
  expect_equal(ds_$text, c("new 1", "paragraph 1"))
  doc <- read_docx()
  expect_warning(
    body_remove(doc),
    "There is nothing left to remove in the document"
  )
})

test_that("cursor and replacement", {
  doc <- read_docx()
  doc <- body_add_par(doc, "blah blah blah")
  doc <- body_add_par(doc, "blah blah blah")
  doc <- body_add_par(doc, "blah blah blah")
  doc <- body_add_par(doc, "Hello text to replace")
  doc <- body_add_par(doc, "blah blah blah")
  doc <- body_add_par(doc, "blah blah blah")
  doc <- body_add_par(doc, "blah blah blah")
  doc <- body_add_par(doc, "Hello text to replace")
  doc <- body_add_par(doc, "blah blah blah")
  template_file <- print(
    x = doc,
    target = tempfile(fileext = ".docx")
  )

  doc <- read_docx(path = template_file)
  while (cursor_reach_test(doc, "to replace")) {
    doc <- cursor_reach(doc, "to replace")

    doc <- body_add_fpar(
      x = doc,
      pos = "on",
      value = fpar(
        "Here is a link: ",
        hyperlink_ftext(
          text = "yopyop",
          href = "https://cran.r-project.org/"
        )
      )
    )
  }

  doc <- cursor_end(doc)
  doc <- body_add_par(doc, "Yap yap yap yap...")
  expect_equal(
    xml_text(xml_find_all(docx_body_xml(doc), "//w:p")),
    c(
      "blah blah blah",
      "blah blah blah",
      "blah blah blah",
      "Here is a link: yopyop",
      "blah blah blah",
      "blah blah blah",
      "blah blah blah",
      "Here is a link: yopyop",
      "blah blah blah",
      "Yap yap yap yap..."
    )
  )
})

to_docx <- function(docx, value, title, subtitle) {
  # Add table
  pre_label <- seq_id <- "Table"
  docx <- body_add_table(docx, value = value, align_table = "center")

  # Add title above table
  run_num <- run_autonum(seq_id = seq_id, pre_label = pre_label)
  title <- block_caption(title, style = "Image Caption", autonum = run_num)
  docx <- body_add_caption(docx, title, pos = "before")
  # Add subtitle above table
  subtitle <- fpar(ftext(subtitle, prop = fp_text(color = "red")))
  docx <- body_add_fpar(docx, subtitle, style = "Normal", pos = "after")
  docx <- cursor_end(docx)

  invisible(docx)
}


test_that("cursor before table", {
  docx <- read_docx()
  tab <- head(mtcars[1:3])

  docx <- to_docx(docx, tab, "Title 1", "Subtitle 1")
  docx <- to_docx(docx, tab, "Title 2", "Subtitle 2")

  nodes_body <- xml_find_all(docx_body_xml(docx), "//w:body/*")

  expect_equal(
    xml_name(nodes_body),
    c("p", "p", "tbl", "p", "p", "tbl", "sectPr")
  )
})

test_that("autoHyphenation", {
  x <- read_docx()
  x$settings$auto_hyphenation <- TRUE
  new_file <- print(x, target = tempfile(fileext = ".docx"))
  new_dir <- unpack_folder(file = new_file, folder = tempfile(pattern = "dir"))
  xml_settings_file <- file.path(new_dir, "word/settings.xml")
  xml_settings <- read_xml(xml_settings_file)
  autoHyphenation_node <- xml_find_first(xml_settings, "//w:autoHyphenation")
  expect_false(inherits(autoHyphenation_node, "xml_missing"))
})
