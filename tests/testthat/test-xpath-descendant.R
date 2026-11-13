# Queries on the body of documents are written `/descendant::x[...]` and
# not `//x[...]`: libxml2 evaluates `//x[...]` by building the set of all the
# nodes of the document, limited to 10 million nodes (error "growing nodeset
# hit limit", or R crashes, on large documents, flextable #737). Both forms
# select the same nodes when predicates are on attributes, as here.

rich_docx <- function() {
  img <- file.path(R.home("doc"), "html", "logo.jpg")
  doc <- read_docx()
  doc <- body_add_img(doc, img, width = 1, height = 1)
  doc <- body_add_fpar(doc, fpar(
    hyperlink_ftext(href = "https://example.org", text = "a link"),
    run_bookmark("bkm", ftext("bookmarked")),
    run_comment(cmt = block_list(fpar("comment")), run = ftext("commented"))
  ))
  doc <- body_add_par(doc, "titre", style = "heading 1")
  doc <- body_add_docx(doc, system.file("doc_examples", "example.docx", package = "officer"))
  doc
}

same_nodes <- function(doc, old, new, ns = xml2::xml_ns(doc)) {
  a <- xml2::xml_find_all(doc, old, ns = ns)
  b <- xml2::xml_find_all(doc, new, ns = ns)
  expect_gt(length(a), 0)
  expect_identical(xml2::xml_path(a), xml2::xml_path(b), label = new)
}

test_that("/descendant:: queries select the nodes of // queries", {
  doc <- rich_docx()
  body <- doc$doc_obj$get()
  ns <- c(
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
  )
  same_nodes(body, "//w:altChunk[@r:id]", "/descendant::w:altChunk[@r:id]", ns)
  same_nodes(body, "//a:blip[@r:embed]", "/descendant::a:blip[@r:embed]", ns)
  same_nodes(body, "//w:hyperlink[@r:id]", "/descendant::w:hyperlink[@r:id]", ns)
  same_nodes(body, "//w:bookmarkStart[@w:name]", "/descendant::w:bookmarkStart[@w:name]", ns)
  same_nodes(
    body,
    "//w:bookmarkStart[@w:name='bkm']/following-sibling::w:r",
    "/descendant::w:bookmarkStart[@w:name='bkm']/following-sibling::w:r",
    ns
  )
  same_nodes(
    body,
    "//w:pStyle[@w:val='Titre1']",
    "/descendant::w:pStyle[@w:val='Titre1']",
    ns
  )
  same_nodes(
    body,
    "/w:document/w:body//*[self::w:commentRangeStart]",
    "/w:document/w:body/descendant::w:commentRangeStart",
    ns
  )
})

test_that("documents with images, links, bookmarks and imported content are written", {
  path <- print(rich_docx(), target = tempfile(fileext = ".docx"))
  x <- read_docx(path)
  body <- x$doc_obj$get()
  ns <- c(
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
  )
  # after print(), images and imported documents are relationships
  same_nodes(
    body,
    "//a:blip[contains(@r:embed, 'rId')]",
    "/descendant::a:blip[contains(@r:embed, 'rId')]",
    ns
  )
  chunk <- xml2::xml_find_all(body, "/descendant::w:altChunk[@r:id]", ns)
  expect_length(chunk, 1)
  expect_match(xml2::xml_attr(chunk, "id"), "^rId[0-9]+$")
  expect_gt(length(xml2::xml_find_all(body, "/descendant::w:hyperlink[@r:id]")), 0)
  expect_identical(nrow(docx_comments(x)), 1L)
  expect_true("bkm" %in% docx_bookmarks(x))
})
