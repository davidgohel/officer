test_that("styles_info is returning a df", {
  x <- read_docx()
  df <- styles_info(x)

  expect_is(df, "data.frame")
  expect_true(all(
    c("style_type", "style_id", "style_name", "is_custom", "is_default") %in%
      names(df)
  ))
  expect_is(df$style_type, "character")
  expect_is(df$style_id, "character")
  expect_is(df$style_name, "character")
  expect_is(df$is_custom, "logical")
  expect_is(df$is_default, "logical")
})

test_that("set par style", {
  doc <- read_docx()
  doc <- docx_set_paragraph_style(
    doc,
    style_id = "rightaligned",
    style_name = "Explicit label",
    fp_p = fp_par(text.align = "right", padding = 20),
    fp_t = fp_text_lite(
      bold = TRUE,
      shading.color = "#FD34F0",
      color = "white"
    )
  )
  doc <- body_add_par(
    doc,
    value = "This is a test",
    style = "Explicit label"
  )

  current_node <- docx_current_block_xml(doc)
  expect_equal(
    xml_attr(xml_child(current_node, "w:pPr/w:pStyle"), "val"),
    "rightaligned"
  )

  df <- styles_info(doc, type = "paragraph")
  df <- df[df$style_id %in% "rightaligned", ]
  expect_true(nrow(df) == 1L)
  if (nrow(df) == 1L) {
    expect_equal(df$shading.color, "FD34F0")
    expect_equal(df$color, "FFFFFF")
    expect_equal(df$align, "right")
  }
})

test_that("set char style", {
  doc <- read_docx()
  doc <- docx_set_character_style(
    doc,
    style_id = "bold",
    base_on = "Default Paragraph Font",
    style_name = "Bold test",
    fp_t = fp_text_lite(bold = TRUE)
  )
  doc <- body_add_fpar(
    doc,
    fpar(run_wordtext(text = "coco", style_id = "bold")),
    style = "Normal"
  )

  current_node <- docx_current_block_xml(doc)
  expect_equal(
    xml_attr(xml_child(current_node, "w:pPr/w:pStyle"), "val"),
    "Normal"
  )

  df <- styles_info(doc, type = "character")
  df <- df[df$style_id %in% "bold", ]
  expect_true(nrow(df) == 1L)
  if (nrow(df) == 1L) {
    expect_equal(df$bold, "true")
    expect_equal(df$style_name, "Bold test")
  }
})

test_that("change_styles changes paragraph styles", {
  # Create a document with different paragraph styles
  doc <- read_docx()
  doc <- body_add_par(doc, "Paragraph with Normal style", style = "Normal")
  doc <- body_add_par(doc, "Paragraph with heading 2", style = "heading 2")
  doc <- body_add_par(doc, "Another Normal paragraph", style = "Normal")

  # Save and reload to ensure styles are properly set
  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  doc <- read_docx(path = tf)

  # Verify original styles
  summary_before <- docx_summary(doc)
  expect_equal(
    summary_before$style_name,
    c("Normal", "heading 2", "Normal")
  )

  # Change styles: Normal and heading 2 -> centered
  mapstyles <- list("centered" = c("Normal", "heading 2"))
  doc <- change_styles(doc, mapstyles = mapstyles)

  # Save and reload to verify changes
  tf2 <- tempfile(fileext = ".docx")
  print(doc, target = tf2)
  doc_changed <- read_docx(path = tf2)

  summary_after <- docx_summary(doc_changed)
  expect_equal(
    summary_after$style_name,
    c("centered", "centered", "centered")
  )

  unlink(tf)
  unlink(tf2)
})

test_that("change_styles changes character styles", {
  # Create a document with character styles
  doc <- read_docx()
  doc <- docx_set_character_style(
    doc,
    style_id = "mystyle",
    base_on = "Default Paragraph Font",
    style_name = "My Style",
    fp_t = fp_text_lite(bold = TRUE)
  )
  doc <- body_add_fpar(
    doc,
    fpar(
      run_wordtext(text = "default text", style_id = "mystyle"),
      ftext(" and styled text", prop = fp_text_lite(bold = TRUE))
    ),
    style = "Normal"
  )

  # Save and reload
  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  doc <- read_docx(path = tf)

  # Change character style
  mapstyles <- list("strong" = "My Style")
  doc <- change_styles(doc, mapstyles = mapstyles)

  all_nodes <- xml2::xml_find_all(
    docx_body_xml(doc),
    "//w:rStyle[@w:val='strong']"
  )
  expect_true(length(all_nodes) > 0)

  unlink(tf)
})

test_that("change_styles handles multiple mappings", {
  doc <- read_docx()
  doc <- body_add_par(doc, "Normal text", style = "Normal")
  doc <- body_add_par(doc, "Heading 1", style = "heading 1")
  doc <- body_add_par(doc, "Heading 2", style = "heading 2")

  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  doc <- read_docx(path = tf)

  # Map multiple styles
  mapstyles <- list(
    "centered" = c("Normal", "heading 2"),
    "heading 1" = "heading 1"
  )
  doc <- change_styles(doc, mapstyles = mapstyles)

  tf2 <- tempfile(fileext = ".docx")
  print(doc, target = tf2)
  doc_changed <- read_docx(path = tf2)

  summary_after <- docx_summary(doc_changed)
  expect_equal(
    summary_after$style_name,
    c("centered", "heading 1", "centered")
  )

  unlink(tf)
  unlink(tf2)
})

test_that("change_styles returns unchanged doc when mapstyles is NULL", {
  doc <- read_docx()
  doc <- body_add_par(doc, "Test paragraph", style = "Normal")

  doc_unchanged <- change_styles(doc, mapstyles = NULL)
  expect_identical(doc, doc_unchanged)

  doc_unchanged2 <- change_styles(doc, mapstyles = list())
  expect_identical(doc, doc_unchanged2)
})

test_that("change_styles errors when style not found", {
  doc <- read_docx()
  doc <- body_add_par(doc, "Test", style = "Normal")

  # Error when 'from' style doesn't exist
  mapstyles <- list("centered" = "NonExistentStyle")
  expect_error(
    change_styles(doc, mapstyles = mapstyles),
    "could not find style.*NonExistentStyle"
  )

  # Error when 'to' style doesn't exist
  mapstyles <- list("NonExistentTarget" = "Normal")
  expect_error(
    change_styles(doc, mapstyles = mapstyles),
    "could not find style.*NonExistentTarget"
  )
})
