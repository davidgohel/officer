test_that("rtf_doc() registers Normal and heading 1..3 with outline levels", {
  doc <- rtf_doc()
  info <- rtf_styles_info(doc)
  expect_equal(
    info$style_id,
    c("Normal", "heading 1", "heading 2", "heading 3")
  )
  expect_equal(info$rtf_index, 1:4)
  expect_equal(info$based_on, c(NA_character_, "Normal", "Normal", "Normal"))
  # Outline levels are stored 0-indexed in the rtf fragment (Heading 1 -> 0).
  expect_match(info$rtf[info$style_id == "heading 1"], "\\\\outlinelevel0")
  expect_match(info$rtf[info$style_id == "heading 2"], "\\\\outlinelevel1")
  expect_match(info$rtf[info$style_id == "heading 3"], "\\\\outlinelevel2")
})

test_that("rtf_styles_info() requires an rtf object", {
  expect_error(rtf_styles_info(list()), "inherits")
})

test_that("rtf_set_paragraph_style() inserts a new style and assigns rtf_index", {
  doc <- rtf_doc()
  before <- max(rtf_styles_info(doc)$rtf_index)
  doc <- rtf_set_paragraph_style(
    doc,
    style_id = "Callout",
    fp_p = fp_par(text.align = "center"),
    fp_t = fp_text_lite(bold = TRUE)
  )
  info <- rtf_styles_info(doc)
  row <- info[info$style_id == "Callout", , drop = FALSE]
  expect_equal(nrow(row), 1L)
  expect_equal(row$rtf_index, before + 1L)
  expect_equal(row$style_name, "Callout")
  expect_equal(row$based_on, "Normal")
})

test_that("rtf_set_paragraph_style() updates an existing style by style_id", {
  doc <- rtf_doc()
  original_index <- rtf_styles_info(doc)$rtf_index[
    rtf_styles_info(doc)$style_id == "heading 1"
  ]
  doc <- rtf_set_paragraph_style(
    doc,
    style_id = "heading 1",
    style_name = "Custom H1",
    fp_p = fp_par(text.align = "center"),
    fp_t = fp_text_lite(bold = TRUE, font.size = 30, color = "#000000"),
    outline_level = 1L
  )
  info <- rtf_styles_info(doc)
  # rtf_index is preserved on replace
  row <- info[info$style_id == "heading 1", , drop = FALSE]
  expect_equal(row$rtf_index, original_index)
  expect_equal(row$style_name, "Custom H1")
  # New properties are in the stored fragment
  expect_match(row$rtf, "\\\\fs60") # 30pt -> \fs60
})

test_that("rtf_set_paragraph_style() validates base_on against existing styles", {
  doc <- rtf_doc()
  expect_error(
    rtf_set_paragraph_style(doc, style_id = "X", base_on = "nope"),
    "does not match any existing style_id"
  )
})

test_that("rtf_set_paragraph_style() validates outline_level range", {
  doc <- rtf_doc()
  expect_error(
    rtf_set_paragraph_style(doc, "X", outline_level = 0L),
    "outline_level"
  )
  expect_error(
    rtf_set_paragraph_style(doc, "X", outline_level = 10L),
    "outline_level"
  )
})

test_that("rtf_add() with style references a registered style in the output", {
  doc <- rtf_doc()
  doc <- rtf_add(doc, "Chapter 1", style = "heading 1")
  doc <- rtf_add(doc, "Body text")
  target <- tempfile(fileext = ".rtf")
  print(doc, target = target)
  rtf <- paste(readLines(target, warn = FALSE), collapse = "\n")
  # styled paragraph emits \pard\plain\sN <duplicated style props>
  expect_match(rtf, "\\\\pard\\\\plain\\\\s2[^\n]*Chapter 1\\\\par")
  # unstyled paragraph keeps the historical emission (\pard with ppr keywords)
  expect_match(rtf, "\\\\pard\\\\sl240[^\n]*Body text")
})

test_that("rtf_add() errors on an unknown style id", {
  doc <- rtf_doc()
  expect_error(
    rtf_add(doc, "x", style = "nonexistent"),
    "is not registered"
  )
})

test_that("rtf_add(fpar, style) and rtf_add(block_list, style) carry the style", {
  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(ftext("Title")), style = "heading 2")
  doc <- rtf_add(
    doc,
    block_list(fpar(ftext("A")), fpar(ftext("B"))),
    style = "heading 3"
  )
  expect_equal(attr(doc$content[[1]], "rtf_style_index"), 3L) # heading 2 -> rtf_index 3
  expect_equal(attr(doc$content[[2]], "rtf_style_index"), 4L) # heading 3 -> rtf_index 4
  expect_equal(attr(doc$content[[3]], "rtf_style_index"), 4L)
})

test_that("walk_based_on_chain returns root-to-leaf order", {
  doc <- rtf_doc()
  parents <- setNames(doc$styles$based_on, doc$styles$style_id)
  expect_equal(
    walk_based_on_chain("heading 1", parents),
    c("Normal", "heading 1")
  )
  expect_equal(walk_based_on_chain("Normal", parents), "Normal")
})

test_that("rtf_add(block_toc()) emits a TOC field", {
  doc <- rtf_doc()
  doc <- rtf_add(doc, block_toc(level = 3))
  target <- tempfile(fileext = ".rtf")
  print(doc, target = target)
  rtf <- paste(readLines(target, warn = FALSE), collapse = "\n")
  # Field switches must be double-escaped in the RTF stream so Word's
  # parser doesn't consume them as RTF control words.
  expect_match(rtf, "\\\\fldinst TOC \\\\\\\\o \"1-3\"")
})

test_that("rtf_add(block_toc(style)) emits a style-keyed TOC field", {
  doc <- rtf_doc()
  doc <- rtf_add(doc, block_toc(style = "heading 1"))
  target <- tempfile(fileext = ".rtf")
  print(doc, target = target)
  rtf <- paste(readLines(target, warn = FALSE), collapse = "\n")
  expect_match(rtf, "\\\\fldinst TOC \\\\\\\\h \\\\\\\\z \\\\\\\\t \"heading 1\"")
})

test_that("walk_based_on_chain terminates on a cycle", {
  # Build a 3-style cycle: A -> B -> C -> A
  doc <- rtf_doc()
  doc <- rtf_set_paragraph_style(doc, "A", base_on = "Normal")
  doc <- rtf_set_paragraph_style(doc, "B", base_on = "A")
  doc <- rtf_set_paragraph_style(doc, "C", base_on = "B")
  # Force a cycle by repointing A onto C
  doc$styles$based_on[doc$styles$style_id == "A"] <- "C"
  parents <- setNames(doc$styles$based_on, doc$styles$style_id)
  chain <- walk_based_on_chain("A", parents)
  # A -> C -> B -> A (cycle): once A is hit again, recursion stops.
  # Normal is not reachable any more since A's parent was repointed.
  expect_setequal(chain, c("A", "B", "C"))
  # Each id appears exactly once: no infinite loop, no duplicates.
  expect_equal(length(chain), length(unique(chain)))
})
