test_that("add notesMaster", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  add_notesMaster(doc)

  doc$content_type$save()
  content_type <- read_xml(file.path(doc$package_dir, "[Content_Types].xml"))
  nodeset <- xml_find_all(content_type, "//d1:Override")
  ct_df <- data.frame(PartName = xml_attr(nodeset, "PartName"), ContentType = xml_attr(nodeset, "ContentType"))
  notesMasterFile <- ct_df$PartName[ct_df$ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"]
  expect_length(notesMasterFile, 1)
  expect_true(file.exists(file.path(doc$package_dir, notesMasterFile)))
  expect_equal(basename(notesMasterFile), names(doc$notesMaster$names()))

  nslideMaster <- doc$notesMaster$collection_get(basename(notesMasterFile))
  rels <- nslideMaster$rel_df()
  theme_file <- rels$target[rels$type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"]
  expect_length(theme_file, 1)
  expect_true(file.exists(file.path(doc$package_dir, dirname(notesMasterFile), theme_file)))
  expect_true(any(grepl(basename(theme_file), ct_df$PartName[ct_df$ContentType == "application/vnd.openxmlformats-officedocument.theme+xml"])))

  xml <- doc$presentation$get()
  node <- xml_find_all(xml, "//p:notesMasterIdLst/p:notesMasterId")
  expect_length(node, 1)
  rid <- xml_attr(node, "id")
  pr_rel <- doc$presentation$rel_df()
  pr_rel <- pr_rel[pr_rel$id == rid, ]
  expect_equal(pr_rel$type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster")
  expect_equal(file.path("/ppt", pr_rel$target), notesMasterFile)
})


test_that("add notesSlide", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  get_or_create_notesSlide(doc)

  doc$content_type$save()
  content_type <- read_xml(file.path(doc$package_dir, "[Content_Types].xml"))
  nodeset <- xml_find_all(content_type, "//d1:Override")
  ct_df <- data.frame(PartName = xml_attr(nodeset, "PartName"), ContentType = xml_attr(nodeset, "ContentType"))
  notesSlideFile <- ct_df$PartName[ct_df$ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"]
  expect_length(notesSlideFile, 1)
  expect_true(file.exists(file.path(doc$package_dir, notesSlideFile)))

  current_slide <- doc$slide$get_slide(doc$cursor)
  rel_df <- current_slide$rel_df()
  notes_ref <- rel_df$target[rel_df$type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"]
  expect_length(notes_ref, 1)
  expect_equal(basename(notes_ref), basename(notesSlideFile))

  idx <- doc$notesSlide$slide_index(basename(notesSlideFile))
  nslide <- doc$notesSlide$get_slide(idx)
  rel_df <- nslide$rel_df()
  master_rel <- rel_df$target[rel_df$type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"]
  expect_equal(basename(master_rel), names(doc$notesMaster$names()))
  slide_rel <- rel_df$target[rel_df$type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"]
  expect_equal(basename(slide_rel), current_slide$name())
})

test_that("add notes to notesSlide", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- set_notes(doc, "I am a test!", notes_location_label("Notes Placeholder 4"))

  nslide <- doc$notesSlide$get_slide(1)
  xml <- nslide$get()
  xpath_ <- "//p:sp[p:nvSpPr/p:cNvPr/@name='Notes Placeholder 4']"
  node_ <- xml_find_all(xml, xpath_)
  expect_false(inherits(node_, "xml_missing"))
  expect_equal(xml_text(node_), "I am a test!")
})

test_that("add notes to notesSlide", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- set_notes(doc, "I am a test!", notes_location_type("body"))

  nslide <- doc$notesSlide$get_slide(1)
  xml <- nslide$get()
  xpath_ <- "//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body']"
  node_ <- xml_find_all(xml, xpath_)
  expect_false(inherits(node_, "xml_missing"))
  expect_equal(xml_text(node_), "I am a test!")
})
test_that("add block_list to notesSlide", {
  value <- block_list(
    fpar(ftext("hello world")),
    fpar(
      ftext("hello"), " ",
      ftext("world")
    ),
    fpar(
      ftext("blah blah blah")
    )
  )

  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- set_notes(doc, value, notes_location_type("body"))

  nslide <- doc$notesSlide$get_slide(1)
  xml <- nslide$get()
  xpath_ <- "//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body']/p:txBody/a:p"
  nodes <- xml_find_all(xml, xpath_)
  expect_true(length(nodes) > 0)
  expect_equal(xml_text(nodes), c("hello world", "hello world", "blah blah blah"))
})

test_that("replace notes on notesSlide", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- set_notes(doc, "I am a test!", notes_location_type("body"))
  doc <- set_notes(doc, "I am a new test!", notes_location_type("body"))

  nslide <- doc$notesSlide$get_slide(1)
  xml <- nslide$get()
  xpath_ <- "//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body']"
  node_ <- xml_find_all(xml, xpath_)
  expect_false(inherits(node_, "xml_missing"))
  expect_equal(xml_text(node_), "I am a new test!")
})
