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

test_that("list markers are resolved in footnotes and comments", {
  doc <- read_docx()
  doc <- body_add(doc, block_list_items(list_item(fpar("Body item"))))
  doc <- body_add_fpar(
    doc,
    fpar(
      "Text",
      run_footnote(
        x = block_list(
          block_list_items(list_item(fpar("Note item")), list_type = "decimal")
        )
      )
    )
  )
  doc <- body_add_fpar(
    doc,
    fpar(
      run_comment(
        cmt = block_list(block_list_items(list_item(fpar("Comment item")))),
        run = ftext("anchor"),
        author = "someone",
        date = "2026-10-08"
      )
    )
  )
  out <- print(doc, target = tempfile(fileext = ".docx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)

  read_part <- function(part) {
    paste(
      readLines(file.path(unpack_dir, "word", part), warn = FALSE),
      collapse = ""
    )
  }
  num_ids <- function(xml) {
    unique(regmatches(xml, gregexpr("w:numId w:val=\"[^\"]+\"", xml))[[1]])
  }

  num_xml <- read_part("numbering.xml")

  for (part in c("document.xml", "footnotes.xml", "comments.xml")) {
    part_xml <- read_part(part)
    # no marker is left behind
    expect_false(grepl("officer-list-", part_xml), label = part)
    ids <- num_ids(part_xml)
    expect_length(ids, 1L)
    id <- gsub("w:numId w:val=\"|\"", "", ids)
    # a w:numId must be an integer, and must be declared in numbering.xml
    expect_match(id, "^[0-9]+$")
    expect_true(
      grepl(sprintf("w:numId=\"%s\"", id), num_xml),
      label = paste(part, "numId", id, "declared")
    )
  }

  # each part gets its own list definition
  all_ids <- unlist(lapply(
    c("document.xml", "footnotes.xml", "comments.xml"),
    function(part) num_ids(read_part(part))
  ))
  expect_length(unique(all_ids), 3L)
})

test_that("adding a list keeps the numbering definitions valid", {
  doc <- read_docx()
  doc <- body_add_par(doc, "Introduction", style = "heading 1")
  doc <- body_add_list(doc, block_list_items(list_item(fpar("Alpha"))))
  doc <- body_add_par(doc, "Methods", style = "heading 1")
  out <- print(doc, target = tempfile(fileext = ".docx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  num_xml <- read_xml(file.path(unpack_dir, "word", "numbering.xml"))
  children <- xml_name(xml_children(num_xml))

  # `CT_Numbering` is a sequence: no `w:abstractNum` may follow a `w:num`,
  # otherwise Word discards the numbering definitions of the document
  expect_true("abstractNum" %in% children)
  expect_true("num" %in% children)
  expect_lt(
    max(which(children == "abstractNum")),
    min(which(children == "num"))
  )
})

test_that("adding a list does not change the numbering of heading styles", {
  # numId of the heading style, and the styles of its list definition
  heading_numbering <- function(add_list) {
    doc <- read_docx()
    doc <- body_add_par(doc, "Introduction", style = "heading 1")
    if (add_list) {
      doc <- body_add_list(doc, block_list_items(list_item(fpar("Alpha"))))
    }
    doc <- body_add_par(doc, "Methods", style = "heading 1")
    out <- print(doc, target = tempfile(fileext = ".docx"))
    unpack_dir <- tempfile()
    unpack_folder(out, unpack_dir)

    styles <- read_xml(file.path(unpack_dir, "word", "styles.xml"))
    numbering <- read_xml(file.path(unpack_dir, "word", "numbering.xml"))
    body <- read_xml(file.path(unpack_dir, "word", "document.xml"))
    style <- xml_find_first(styles, "//w:style[w:name/@w:val='heading 1']")
    num_id <- xml_attr(xml_find_first(style, ".//w:numId"), "val")
    abstract_id <- function(id) {
      xml_attr(
        xml_find_first(
          numbering,
          sprintf("/w:numbering/w:num[@w:numId='%s']/w:abstractNumId", id)
        ),
        "val"
      )
    }
    list(
      style_id = xml_attr(style, "styleId"),
      num_id = num_id,
      abstract_id = abstract_id(num_id),
      abstract_styles = xml_attr(
        xml_find_all(
          numbering,
          sprintf(
            "/w:numbering/w:abstractNum[@w:abstractNumId='%s']//w:pStyle",
            abstract_id(num_id)
          )
        ),
        "val"
      ),
      list_num_id = xml_attr(xml_find_all(body, "//w:numId"), "val"),
      list_abstract_id = vapply(
        xml_attr(xml_find_all(body, "//w:numId"), "val"),
        abstract_id,
        character(1),
        USE.NAMES = FALSE
      )
    )
  }

  reference <- heading_numbering(FALSE)
  expect_false(is.na(reference$num_id))
  expect_true(reference$style_id %in% reference$abstract_styles)

  listed <- heading_numbering(TRUE)
  expect_equal(listed$num_id, reference$num_id)
  expect_equal(listed$abstract_id, reference$abstract_id)
  expect_equal(listed$abstract_styles, reference$abstract_styles)

  # the list has its own definition, the headings do not share it
  expect_length(listed$list_num_id, 1L)
  expect_false(listed$list_num_id == listed$num_id)
  expect_false(is.na(listed$list_abstract_id))
  expect_false(listed$list_abstract_id == listed$abstract_id)
})

test_that("new numbering definitions are inserted before w:numIdMacAtCleanup", {
  ns_w <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  package_dir <- tempfile()
  dir.create(file.path(package_dir, "word"), recursive = TRUE)
  writeLines(
    sprintf(
      paste0(
        "<w:numbering xmlns:w=\"%s\">",
        "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"/>",
        "</w:abstractNum>",
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>",
        "<w:numIdMacAtCleanup w:val=\"5\"/>",
        "</w:numbering>"
      ),
      ns_w
    ),
    file.path(package_dir, "word", "numbering.xml")
  )

  marker <- "officer-list-bullet-aaaaaaaa-0000-0000-0000-000000000000"
  xml_str <- sprintf(
    "<w:p><w:pPr><w:numPr><w:numId w:val=\"%s\"/></w:numPr></w:pPr></w:p>",
    marker
  )
  res <- process_list_markers(xml_str, package_dir)

  expect_false(grepl("officer-list-", res))

  num_xml <- read_xml(file.path(package_dir, "word", "numbering.xml"))
  expect_equal(
    xml_name(xml_children(num_xml)),
    c("abstractNum", "abstractNum", "num", "num", "numIdMacAtCleanup")
  )
})
