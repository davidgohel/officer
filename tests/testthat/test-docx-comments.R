test_that("add comments", {
  fp_bold <- fp_text_lite(bold = TRUE)
  fp_red <- fp_text_lite(color = "red")

  bl <- block_list(
    fpar(ftext("Comment multiple words.", fp_bold)),
    fpar(
      ftext("Second line.", fp_red)
    )
  )

  a_par <- fpar(
    "This paragraph contains",
    run_comment(
      cmt = bl,
      run = ftext("a comment."),
      author = "Author Me",
      date = "2023-06-01"
    )
  )

  b_par <- fpar(
    run_comment(
      cmt = block_list(
        fpar(
          ftext("Comment a paragraph.")
        )
      ),
      run = ftext("This paragraph contains another comment."),
      author = "Author You",
      date = "2023-06-01",
      initials = "OM"
    )
  )
  doc <- read_docx()
  doc <- body_add_fpar(doc, value = a_par, style = "Normal")
  doc <- body_add_fpar(doc, value = b_par, style = "Normal")

  docx_file <- print(doc, target = tempfile(fileext = ".docx"))
  docx_dir <- tempfile()
  unpack_folder(docx_file, docx_dir)

  doc <- read_xml(file.path(docx_dir, "word/comments.xml"))
  comment1 <- xml_find_first(doc, "w:comment[@w:id='0']")
  comment2 <- xml_find_first(doc, "w:comment[@w:id='1']")

  expect_false(inherits(comment1, "xml_missing"))
  expect_false(inherits(comment2, "xml_missing"))

  expect_length(xml_children(comment1), 2)
  expect_length(xml_children(comment2), 1)
})
