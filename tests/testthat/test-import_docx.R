test_that("body_import_docx messages", {
  tf <- tempfile(fileext = ".docx")

  x <- read_docx()
  testthat::expect_warning(
    x <- body_import_docx(
      x = x,
      src = system.file(
        package = "officer", "doc_examples", "example.docx"
      )
    ),
    "Style(s) mapping(s) for 'paragraphs' are missing in the body of the document:",
    fixed = TRUE
  )

  x <- read_docx()
  testthat::expect_no_message(
    x <- body_import_docx(
      x = x,
      src = system.file(
        package = "officer", "doc_examples", "example.docx"
      ),
      par_style_mapping = list(
        "Normal" = c("List Paragraph")
      ),
      tbl_style_mapping = list(
        "Normal Table" = "Light Shading"
      )
    )
  )
  docx_summary(x)
})

test_that("body_import_docx copy the content", {
  tf <- tempfile(fileext = ".docx")
  fi <- system.file(
    package = "officer", "doc_examples", "example.docx"
  )
  x <- read_docx()
  x <- body_import_docx(
    x = x,
    src = fi,
    par_style_mapping = list(
      "Normal" = c("List Paragraph")
    ),
    tbl_style_mapping = list(
      "Normal Table" = "Light Shading"
    )
  )
  print(x, target = tf)

  old_doc <- read_docx(fi)
  old_doc_head <- docx_summary(old_doc)

  new_doc <- read_docx(tf)
  new_doc_head <- docx_summary(new_doc)

  expect_equal(
    head(new_doc_head)[c("doc_index", "text")],
    head(old_doc_head)[c("doc_index", "text")]
  )
  expect_equal(
    head(new_doc_head)$style_name,
    c("heading 1", NA, "heading 1", "Normal", "Normal", "Normal")
  )

})

test_that("body_import_docx preserves abstractNum and num correspondence", {
  tf <- tempfile(fileext = ".docx")
  template_file <- "docs_dir/officer-lists.docx"
  # Create front document
  front <- read_docx()
  front <- body_add_par(front, "Frontmatter", style = "heading 1")
  tf_front <- tempfile(fileext = ".docx")
  print(front, target = tf_front)

  # Merge with document containing lists
  merged <- read_docx(tf_front)
  merged <- cursor_end(merged)
  merged <- body_import_docx(
    merged,
    par_style_mapping = list(
      "Normal" = c("First Paragraph", "Compact")
    ),
    src = template_file
  )
  print(merged, target = tf)

  # Read the front document and extract numbering information
  front_doc <- read_docx(tf_front)
  numbering_file <- file.path(front_doc$package_dir, "word/numbering.xml")
  front_numbering_xml <- xml2::read_xml(numbering_file)
  # Read the list document and extract numbering information
  list_doc <- read_docx(template_file)
  numbering_file <- file.path(list_doc$package_dir, "word/numbering.xml")
  list_numbering_xml <- xml2::read_xml(numbering_file)
  # Read the merged document and extract numbering information
  merged_doc <- read_docx(tf)
  numbering_file <- file.path(merged_doc$package_dir, "word/numbering.xml")
  merged_numbering_xml <- xml2::read_xml(numbering_file)


  # Extract abstractNum IDs
  abstractnum_nodes <- xml2::xml_find_all(merged_numbering_xml, "//w:abstractNum")
  abstractnum_ids <- xml2::xml_attr(abstractnum_nodes, "abstractNumId")

  # Extract num elements and their associated abstractNumId references
  num_nodes <- xml2::xml_find_all(merged_numbering_xml, "//w:num")
  num_ids <- xml2::xml_attr(num_nodes, "numId")

  # For each num, get the referenced abstractNumId
  abstractnum_refs <- sapply(num_nodes, function(node) {
    ref_node <- xml2::xml_find_first(node, "w:abstractNumId")
    xml2::xml_attr(ref_node, "val")
  })

  # Verify that all referenced abstractNumIds exist in the document
  expect_true(
    all(abstractnum_refs %in% abstractnum_ids),
    info = "All num elements should reference existing abstractNum definitions"
  )

  # Verify that abstractNum IDs are unique
  expect_equal(
    length(abstractnum_ids),
    length(unique(abstractnum_ids)),
    info = "abstractNum IDs should be unique"
  )

  # Extract abstractNum IDs from front document
  front_abstractnum_nodes <- xml2::xml_find_all(
    front_numbering_xml,
    "//w:abstractNum"
  )
  front_abstractnum_ids <- xml2::xml_attr(
    front_abstractnum_nodes,
    "abstractNumId"
  )

  # Extract num IDs from front document
  front_num_nodes <- xml2::xml_find_all(front_numbering_xml, "//w:num")
  front_num_ids <- xml2::xml_attr(front_num_nodes, "numId")

  # Extract abstractNum IDs from list document
  list_abstractnum_nodes <- xml2::xml_find_all(
    list_numbering_xml,
    "//w:abstractNum"
  )
  list_abstractnum_ids <- xml2::xml_attr(list_abstractnum_nodes, "abstractNumId")

  # Extract num IDs from list document
  list_num_nodes <- xml2::xml_find_all(list_numbering_xml, "//w:num")
  list_num_ids <- xml2::xml_attr(list_num_nodes, "numId")

  # Verify that the number of abstractNum in merged document equals
  # the sum from front and list documents
  expect_equal(
    length(abstractnum_ids),
    length(front_abstractnum_ids) + length(list_abstractnum_ids),
    info = paste(
      "Merged document should have",
      length(front_abstractnum_ids) + length(list_abstractnum_ids),
      "abstractNum elements"
    )
  )

  # Verify that the number of num in merged document equals
  # the sum from front and list documents
  expect_equal(
    length(num_ids),
    length(front_num_ids) + length(list_num_ids),
    info = paste(
      "Merged document should have",
      length(front_num_ids) + length(list_num_ids),
      "num elements"
    )
  )

  # Verify that abstractNum IDs in merged document are sequential iterations
  # Convert IDs to integers
  front_abstractnum_ids_int <- as.integer(front_abstractnum_ids)
  list_abstractnum_ids_int <- as.integer(list_abstractnum_ids)
  merged_abstractnum_ids_int <- as.integer(abstractnum_ids)

  # The merged document should have:
  # - All IDs from front document (unchanged)
  # - All IDs from list document (renumbered to follow front IDs)
  max_front_abstractnum <- if (length(front_abstractnum_ids_int) > 0) {
    max(front_abstractnum_ids_int)
  } else {
    -1L
  }

  # Expected IDs: front IDs + list IDs renumbered starting after max front ID
  expected_abstractnum_ids <- c(
    front_abstractnum_ids_int,
    seq_along(list_abstractnum_ids_int) + max_front_abstractnum
  )

  expect_equal(
    sort(merged_abstractnum_ids_int),
    sort(expected_abstractnum_ids),
    info = "abstractNum IDs should be properly renumbered in merged document"
  )

  unlink(tf_front)
  unlink(tf)
})

