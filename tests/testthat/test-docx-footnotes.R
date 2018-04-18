context("add footnotes in docx")




test_that("add footnotes", {

  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  bold_fp <- shortcuts$fp_bold()
  normal_fp <- fp_text()
  bl1 <- block_list(
    fpar(ftext("hello", normal_fp), ftext(" world!", bold_fp)),
    fpar(
      ftext("Do you enjoy", bold_fp),
      stext(" with ", "strong"),
      external_img(src = img.file, height = 1.06/3, width = 1.39/3)
    )
  )
  bl2 <- block_list(
    fpar(ftext("hello", normal_fp), ftext(" world!", bold_fp))
  )

  x <- read_docx() %>%
    body_add_par("Here is a footnote", style = "Normal") %>%
    slip_in_footnote(style = "reference_id", blocks = bl1) %>%
    body_add_par("Here is another footnote", style = "Normal") %>%
    slip_in_footnote(style = "reference_id", blocks = bl2)

  docx_file <- tempfile(fileext = ".docx")
  docx_dir <- tempfile()
  print(x, target = docx_file)

  unpack_folder(docx_file, docx_dir)

  doc <- read_xml(file.path(docx_dir, "word/footnotes.xml"))
  footnote1 <- xml_find_first(doc, "w:footnote[@w:id='1']")
  footnote2 <- xml_find_first(doc, "w:footnote[@w:id='2']")

  expect_false( inherits(footnote1, "xml_missing") )
  expect_false( inherits(footnote2, "xml_missing") )

  expect_length(xml_children(footnote1), 2)
  expect_length(xml_children(footnote2), 1)

})

test_that("check blocks", {

  x <- read_docx() %>%
    body_add_par("Here is a footnote", style = "Normal")
  expect_error(x %>% slip_in_footnote(style = "reference_id", blocks = "hi"),
               regexp = "block_list")

})


