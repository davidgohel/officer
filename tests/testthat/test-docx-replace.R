context("replace elements in an rdocx")


test_that("docx replace_at bkm", {
  doc <- read_docx() %>%
    body_add_par("centered text", style = "centered") %>%
    slip_in_text(". How are you", style = "strong") %>%
    body_bookmark("text_to_replace")

  xmldoc <- doc$doc_obj$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "text_to_replace")
  bm_start <- xml_find_first(xmldoc, xpath_)
  expect_false( inherits(bm_start, "xml_missing"))

  doc <- body_replace_at(doc, "text_to_replace", "not left aligned")
  xmldoc <- doc$doc_obj$get()
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", "text_to_replace")
  newtext <- xml_find_first(xmldoc, "w:body/w:p[2]/w:r")
  expect_equal(xml_text(newtext), "not left aligned")

  lasttext <- xml_find_first(xmldoc, "w:body/w:p[2]/w:r[2]")
  expect_equal(xml_text(lasttext), ". How are you")
})

