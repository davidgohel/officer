test_that("replace bkm with text in body", {
  doc <- read_docx()
  fp <- fpar(
    run_bookmark(
      ftext(
        "centered text"),
      bkm = "text_to_replace"
    ),
    ftext(
      text = ". How are you",
      prop = fp_text(
        color = NA, font.size = NA, bold = TRUE, italic = NA,
        underlined = NA, font.family = NA_character_,
        cs.family = NA_character_, eastasia.family = NA_character_,
        hansi.family = NA_character_, shading.color = NA_character_)
      )
  )
  doc <- body_add_fpar(doc, value = fp, style = "centered")

  xmldoc <- doc$doc_obj$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "text_to_replace")
  bm_start <- xml_find_first(xmldoc, xpath_)
  expect_false( inherits(bm_start, "xml_missing"))

  doc <- body_replace_text_at_bkm(doc, "text_to_replace", "not left aligned")
  xmldoc <- doc$doc_obj$get()
  newtext <- xml_find_first(xmldoc, "w:body/w:p[1]/w:r")
  expect_equal(xml_text(newtext), "not left aligned")

  lasttext <- xml_find_first(xmldoc, "w:body/w:p[1]/w:r[2]")
  expect_equal(xml_text(lasttext), ". How are you")
})

test_that("replace bkm with images in header/footer", {

  template <- system.file(package = "officer", "doc_examples/example.docx")
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )

  doc <- read_docx(path = template)
  doc <- headers_replace_img_at_bkm(x = doc, bookmark = "bmk_header",
                                    value = external_img(src = img.file, width = .53, height = .7))
  doc <- footers_replace_img_at_bkm(x = doc, bookmark = "bmk_footer",
                                    value = external_img(src = img.file, width = .53, height = .7))
  print(doc, target = "test_replace_img.docx")

  doc <- read_docx(path = "test_replace_img.docx")

  xmldoc <- doc$headers[[1]]$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "bmk_header")
  bm_start <- xml_find_first(xmldoc, xpath_)
  expect_false( inherits(bm_start, "xml_missing"))

  blip <- xml_find_first(xmldoc, "//w:p/w:r/w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip")
  expect_equal( xml_attr(blip, "embed"), "rId1")

  xmldoc <- doc$footers[[1]]$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "bmk_footer")
  bm_start <- xml_find_first(xmldoc, xpath_)
  expect_false( inherits(bm_start, "xml_missing"))

  blip <- xml_find_first(xmldoc, "//w:p/w:r/w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip")
  expect_equal( xml_attr(blip, "embed"), "rId1")

})

test_that("replace bkm in headers and footers", {

  doc <- read_docx(path = "docs_dir/table-complex.docx" )
  doc <- headers_replace_text_at_bkm(doc, "hello1", "salut")
  doc <- footers_replace_text_at_bkm(doc, "hello2", "salut")

  xmldoc <- doc$headers[[1]]$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "hello1")
  bm_start <- xml_find_first(xmldoc, xpath_)
  expect_false( inherits(bm_start, "xml_missing"))
  newtext <- xml_find_first(xmldoc, "/w:hdr/w:p[1]/w:r")
  expect_equal(xml_text(newtext), "salut")

  xmldoc <- doc$footers[[1]]$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "hello2")
  bm_start <- xml_find_first(xmldoc, xpath_)
  expect_false( inherits(bm_start, "xml_missing"))
  newtext <- xml_find_first(xmldoc, "/w:ftr/w:p[1]/w:r[3]")
  expect_equal(xml_text(newtext), "salut")

})

test_that("docx replace text", {
  doc <- read_docx()
  doc <- body_add_par(doc, "Placeholder one")
  doc <- body_add_par(doc, "Placeholder two")
  doc <- body_replace_all_text(doc, old_value = "placeholder", new_value = "new",
                               only_at_cursor = FALSE, ignore.case = TRUE)
  xmldoc <- doc$doc_obj$get()
  expect_equal(xml_text( xml_find_all(xmldoc, "//w:p") ), c("new one", "new two") )
})

test_that("docx replace all text", {
  doc <- read_docx(path = "docs_dir/table-complex.docx" )

  doc <- headers_replace_all_text(doc, old_value = "hello", new_value = "salut",
                               only_at_cursor = FALSE, ignore.case = TRUE)
  doc <- footers_replace_all_text(doc, old_value = "hello", new_value = "salut",
                               only_at_cursor = FALSE, ignore.case = TRUE)

  xmldoc <- doc$headers[[1]]$get()
  expect_equal(xml_text( xml_find_all(xmldoc, "//w:p") ), "salut world" )

  xmldoc <- doc$footers[[1]]$get()
  expect_equal(xml_text( xml_find_all(xmldoc, "//w:p") ), "world salut" )
})

unlink("*.docx", force = TRUE)
