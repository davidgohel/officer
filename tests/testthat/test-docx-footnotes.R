test_that("add footnotes", {
  fp_bold <- fp_text_lite(bold = TRUE)
  fp_refnote <- fp_text_lite(vertical.align = "superscript")

  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  bl <- block_list(
    fpar(ftext("hello", fp_bold)),
    fpar(
      ftext("hello world", fp_bold),
      external_img(src = img.file, height = 1.06, width = 1.39)
    )
  )

  a_par <- fpar(
    "this paragraph contains a note ",
    run_footnote(x = bl, prop = fp_refnote),
    "."
  )
  b_par <- fpar(
    "this paragraph contains another note ",
    run_footnote(
      x = block_list(
        fpar(ftext("Salut", fp_text_lite()))
      ),
      prop = fp_refnote
    ),
    "."
  )

  doc <- read_docx()
  doc <- body_add_fpar(doc, value = a_par, style = "Normal")
  doc <- body_add_fpar(doc, value = b_par, style = "Normal")

  docx_file <- print(doc, target = tempfile(fileext = ".docx"))
  docx_dir <- tempfile()
  unpack_folder(docx_file, docx_dir)

  doc <- read_xml(file.path(docx_dir, "word/footnotes.xml"))
  footnote1 <- xml_find_first(doc, "w:footnote[@w:id='1']")
  footnote2 <- xml_find_first(doc, "w:footnote[@w:id='2']")

  expect_false(inherits(footnote1, "xml_missing"))
  expect_false(inherits(footnote2, "xml_missing"))

  expect_length(xml_children(footnote1), 2)
  expect_length(xml_children(footnote2), 1)
})
