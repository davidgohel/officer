test_that("docx summary", {
  example_docx <- system.file(package = "officer", "doc_examples/example.docx")

  doc <- read_docx(path = example_docx)
  doc_data <- docx_summary(doc)
  table_data <- subset(doc_data, content_type %in% "table cell" & is_header)
  expect_equal(table_data$text, c("Petals", "Internode", "Sepal", "Bract"))

  doc_data_pr <- docx_summary(doc, preserve = TRUE)
  expect_equal(doc_data_pr[28, ][["text"]], "Note\nNew line note")

  doc <- read_docx()
  doc <- body_add_fpar(
    x = doc,
    value = fpar(run_word_field(field = "Date \\@ \"MMMM d yyyy\""))
  )
  x <- print(doc, tempfile(fileext = ".docx"))
  doc <- read_docx(x)
  expect_equal(docx_summary(doc)$text, "Date \\@ \"MMMM d yyyy\"")
  doc <- read_docx(x)
  expect_equal(docx_summary(doc, remove_fields = TRUE)$text, "")
})

test_that("complex docx table", {
  doc <- read_docx(path = "docs_dir/table-complex.docx")
  doc_data <- docx_summary(doc)
  table_data <- doc_data[doc_data$content_type == "table cell", ]
  table_data <- table_data[order(table_data$row_id, table_data$cell_id), ]

  first_row <- table_data[table_data$row_id %in% 1, ]

  expect_equal(
    first_row$text,
    c(
      "Column head",
      "column head",
      "column head",
      "column head",
      NA,
      NA,
      "x",
      "y"
    )
  )
  expect_equal(first_row$col_span, c(1, 1, 1, 3, 0, 0, 1, 1))
  expect_equal(first_row$row_span, c(2, 2, 2, 1, 1, 1, 2, 2))
  expect_true(all(
    table_data$text[table_data$col_span < 1 | table_data$row_span < 1] %in%
      NA_character_
  ))
})

test_that("preserves non breaking hyphens", {
  doc <- read_docx()

  make_xml_elt <- function(x, y) {
    sprintf(
      paste0(
        officer:::wp_ns_yes,
        "<w:pPr><w:pStyle w:val=\"Normal\"/></w:pPr><w:r>",
        "<w:t xml:space=\"preserve\">%s</w:t>",
        "<w:noBreakHyphen/>",
        "<w:t xml:space=\"preserve\">%s</w:t>",
        "</w:r></w:p>"
      ),
      x,
      y
    )
  }
  doc <- officer:::body_add_xml(
    x = doc,
    str = make_xml_elt("Inspector", "General")
  )
  doc <- officer:::body_add_xml(
    x = doc,
    str = make_xml_elt("General", "Inspector")
  )
  expect_equal(
    docx_summary(doc)[["text"]],
    c("Inspector\u002DGeneral", "General\u002DInspector")
  )
})

test_that("detailed summary", {
  doc <- read_docx()

  fpar_ <- fpar(
    ftext("Formatted ", prop = fp_text(bold = TRUE, color = "red")),
    ftext(
      "paragraph ",
      prop = fp_text(
        shading.color = "blue"
      )
    ),
    ftext(
      "with multiple runs.",
      prop = fp_text(italic = TRUE, font.size = 20, font.family = "Arial")
    )
  )

  doc <- body_add_fpar(doc, fpar_, style = "Normal")

  fpar_ <- fpar(
    "Unformatted ",
    "paragraph ",
    "with multiple runs."
  )

  doc <- body_add_fpar(doc, fpar_, style = "Normal")

  doc <- body_add_par(doc, "Single Run", style = "Normal")

  doc <- body_add_fpar(
    doc,
    fpar(
      "Single formatetd run ",
      fp_t = fp_text(bold = TRUE, color = "red")
    )
  )

  xml_elt <- paste0(
    officer:::wp_ns_yes,
    "<w:pPr><w:pStyle w:val=\"Normal\"/></w:pPr>",
    "<w:r><w:rPr></w:rPr><w:t>NA</w:t></w:r>",
    "<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>toggle</w:t></w:r>",
    "<w:r><w:rPr><w:b w:val=\"0\"/><w:i w:val=\"0\"/></w:rPr><w:t>0</w:t></w:r>",
    "<w:r><w:rPr><w:b w:val=\"1\"/><w:i w:val=\"1\"/></w:rPr><w:t>1</w:t></w:r>",
    "<w:r><w:rPr><w:b w:val=\"false\"/><w:i w:val=\"false\"/></w:rPr><w:t>false</w:t></w:r>",
    "<w:r><w:rPr><w:b w:val=\"true\"/><w:i w:val=\"true\"/></w:rPr><w:t>true</w:t></w:r>",
    "<w:r><w:rPr><w:b w:val=\"off\"/><w:i w:val=\"off\"/></w:rPr><w:t>off</w:t></w:r>",
    "<w:r><w:rPr><w:b w:val=\"on\"/><w:i w:val=\"on\"/></w:rPr><w:t>on</w:t></w:r>",
    "</w:p>"
  )

  doc <- officer:::body_add_xml(
    x = doc,
    str = xml_elt
  )

  # #662
  xml_elt <- paste0(
    officer:::wp_ns_yes,
    "<w:pPr><w:pStyle w:val=\"Normal\"/></w:pPr>",
    "<w:r><w:rPr><w:u/><w:sz/><w:szCs/><w:color/><w:shd/></w:rPr><w:t>no attributes</w:t></w:r>",
    "<w:r><w:rPr><w:shd w:val=\"clear\"/></w:rPr><w:t>shading wo color and fill</w:t></w:r>",
    "</w:p>"
  )

  doc <- officer:::body_add_xml(
    x = doc,
    str = xml_elt
  )

  doc_sum <- docx_summary(doc, detailed = TRUE)

  expect_true("run" %in% names(doc_sum))
  expect_type(doc_sum$run, "list")
  expect_equal(lengths(doc_sum$run), rep(12, 6))
  expect_equal(sapply(doc_sum$run, nrow), c(3, 3, 1, 1, 8, 2))

  expect_true(all(sapply(doc_sum$run$bold, is.logical)))
  expect_true(all(sapply(doc_sum$run$italic, is.logical)))
  expect_true(all(sapply(doc_sum$run$sz, is.integer)))
  expect_true(all(sapply(doc_sum$run$szCs, is.integer)))
  expect_true(all(sapply(doc_sum$run$underline, is_character)))
  expect_true(all(sapply(doc_sum$run$color, is_character)))
  expect_true(all(sapply(doc_sum$run$shading, is_character)))
  expect_true(all(sapply(doc_sum$run$shading_color, is_character)))
  expect_true(all(sapply(doc_sum$run$shading_fill, is_character)))
})


test_that("pptx summary", {
  example_pptx <- system.file(package = "officer", "doc_examples/example.pptx")
  doc <- read_pptx(path = example_pptx)
  doc_data <- pptx_summary(doc)
  table_data <- subset(
    doc_data,
    content_type %in% "table cell" & row_id == 1 & slide_id == 1
  )
  expect_equal(table_data$text, c("Header 1 ", "Header 2", "Header 3"))

  table_data <- subset(
    doc_data,
    content_type %in% "table cell" & row_id == 4 & slide_id == 1
  )
  expect_equal(table_data$text, c("B", "9.0", "Salut"))

  example_pptx <- "docs_dir/table-complex.pptx"
  doc <- read_pptx(path = example_pptx)
  doc_data <- pptx_summary(doc)
  table_data <- subset(
    doc_data,
    content_type %in% "table cell" & cell_id == 3 & slide_id == 1
  )
  expect_equal(
    table_data$text,
    c(
      "Header 3",
      NA,
      "blah blah blah",
      "Salut",
      "Hello",
      "sisi",
      NA,
      NA,
      NA,
      NA
    )
  )
  expect_equal(table_data$row_span, c(1, 1, 1, 1, 1, 4, 0, 0, 0, 1))
  expect_equal(table_data$col_span, c(1, 0, 1, 1, 1, 1, 1, 1, 1, 0))

  slide2_data <- subset(doc_data, content_type %in% "paragraph" & slide_id == 2)
  expect_equal(slide2_data$text, c("coco", "line of text", "blah blah blah"))

  doc_data_pr <- pptx_summary(doc, preserve = TRUE)
  expect_equal(doc_data_pr[23, ][["text"]], "blah\n \nblah\n \nblah")
})


test_that("empty slide summary", {
  example_pptx <- "docs_dir/test_empty.pptx"
  doc <- read_pptx(path = example_pptx)
  run <- try(pptx_summary(doc), silent = TRUE)
  expect_false(inherits(run, "try-error"))
})
