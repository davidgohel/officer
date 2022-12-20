test_that("styles_info is returning a df", {
  x <- read_docx()
  df <- styles_info(x)

  expect_is(df, "data.frame")
  expect_true(all(c("style_type", "style_id", "style_name", "is_custom", "is_default") %in% names(df)))
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
      color = "white")
  )
  doc <- body_add_par(
    doc, value = "This is a test",
    style = "Explicit label")

  current_node <- docx_current_block_xml(doc)
  expect_equal(
    xml_attr(xml_child(current_node, "w:pPr/w:pStyle"), "val"),
    "rightaligned")

  df <- styles_info(doc, type = "paragraph")
  df <- df[df$style_id %in% "rightaligned",]
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
    doc, fpar(run_wordtext(text = "coco", style_id = "bold")),
    style = "Normal")

  current_node <- docx_current_block_xml(doc)
  expect_equal(
    xml_attr(xml_child(current_node, "w:pPr/w:pStyle"), "val"),
    "Normal")

  df <- styles_info(doc, type = "character")
  df <- df[df$style_id %in% "bold",]
  expect_true(nrow(df) == 1L)
  if (nrow(df) == 1L) {
    expect_equal(df$bold, "true")
    expect_equal(df$style_name, "Bold test")
  }

})

