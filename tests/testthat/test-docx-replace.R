context("replace elements in an rdocx")

test_that("replace bkm with text in body", {
  doc <- read_docx() %>%
    body_add_par("centered text", style = "centered") %>%
    slip_in_text(". How are you", style = "strong") %>%
    body_bookmark("text_to_replace")

  xmldoc <- doc$doc_obj$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "text_to_replace")
  bm_start <- xml_find_first(xmldoc, xpath_)
  expect_false( inherits(bm_start, "xml_missing"))

  doc <- body_replace_text_at_bkm(doc, "text_to_replace", "not left aligned")
  xmldoc <- doc$doc_obj$get()
  newtext <- xml_find_first(xmldoc, "w:body/w:p[2]/w:r")
  expect_equal(xml_text(newtext), "not left aligned")

  lasttext <- xml_find_first(xmldoc, "w:body/w:p[2]/w:r[2]")
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
  doc <- read_docx() %>%
    body_add_par("Placeholder one") %>%
    body_add_par("Placeholder two")
  doc <- body_replace_all_text(doc, old_value = "placeholder", new_value = "new",
                               only_at_cursor = FALSE, ignore.case = TRUE)
  xmldoc <- doc$doc_obj$get()
  expect_equal(xml_text( xml_find_all(xmldoc, "//w:p") ), c("", "new one", "new two") )
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
