test_that("body_add_table", {
  x <- read_docx()
  x <- body_add_table(x, value = iris, style = "table_template")
  node <- docx_current_block_xml(x)
  expect_equal(xml_name(node), "tbl")
})

test_that("wml structure", {
  x <- read_docx()
  x <- body_add_table(x, value = iris, style = "table_template")
  node <- docx_current_block_xml(x)
  expect_length(xml_find_all(node, xpath = "w:tr[w:trPr/w:tblHeader]"), 1)
  expect_length(
    xml_find_all(node, xpath = "w:tr[not(w:trPr/w:tblHeader)]"),
    nrow(iris)
  )
  expect_length(
    xml_find_all(
      node,
      xpath = "w:tr[not(w:trPr/w:tblHeader)]/w:tc/w:p/w:r/w:t"
    ),
    ncol(iris) * nrow(iris)
  )
})

test_that("block_caption", {
  run_num <- run_autonum(
    seq_id = "tab",
    pre_label = "tab. ",
    bkm = "mtcars_table"
  )

  expect_output(
    print(block_caption("mtcars table", style = "Normal")),
    "caption \\[autonum off\\]: mtcars table"
  )
  expect_output(
    print(block_caption("mtcars table", style = "Normal", autonum = run_num)),
    "caption \\[autonum on\\]: mtcars table"
  )

  caption <- block_caption("mtcars table", autonum = run_num)
  expect_equal(caption$style, "Normal")

  expect_match(
    to_wml(caption, knitting = TRUE),
    "::: \\{custom-style=\"Normal\"\\}"
  )
  expect_match(to_wml(caption, knitting = TRUE), "mtcars table\\n:::")
  expect_match(to_wml(caption, knitting = FALSE), "^<w:p>")
  expect_match(to_wml(caption, knitting = FALSE), "</w:p>$")
  expect_match(to_wml(caption, knitting = FALSE), "mtcars table</w:t>")
})

test_that("names stay as is", {
  df <- data.frame(
    "hello coco" = c(1, 2),
    value = c("одно значение", "еще значение"),
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
  x <- read_docx()
  x <- body_add_table(x, value = df, style = "table_template")
  node <- docx_current_block_xml(x)
  first_text <- xml_find_first(
    node,
    xpath = "w:tr[w:trPr/w:tblHeader]/w:tc/w:p/w:r/w:t"
  )
  first_text <- xml_text(first_text)
  expect_equal(first_text, "hello coco")
})
