context("doc properties")

test_that("read docx properties", {
  doc <- read_docx(path = system.file(package = "officer", "doc_examples/example.docx"))
  properties <- doc_properties(doc)
  expect_equal( properties$value[properties$tag %in% "title"], "document title" )
  expect_equal( properties$value[properties$tag %in% "subject"], "document subject" )
  expect_equal( properties$value[properties$tag %in% "creator"], "author" )
  expect_equal( properties$value[properties$tag %in% "description"], "these are comments" )
  expect_equal( properties$value[properties$tag %in% "created"], "2017-04-26T13:10:00Z" )

})

test_that("read pptx properties", {
  doc <- read_pptx(path = system.file(package = "officer", "doc_examples/example.pptx"))
  properties <- doc_properties(doc)
  expect_equal( properties$value[properties$tag %in% "title"], "document title" )
  expect_equal( properties$value[properties$tag %in% "subject"], "document subject" )
  expect_equal( properties$value[properties$tag %in% "creator"], "author" )
  expect_equal( properties$value[properties$tag %in% "description"], "these are comments" )
  expect_equal( properties$value[properties$tag %in% "created"], "2017-04-27T11:29:40Z" )

})


test_that("set docx properties", {
  doc <- read_docx()
  time_now <- Sys.time()
  filename <- tempfile(fileext = ".docx")
  doc  %>%
    set_doc_properties(title = "title",
                       subject = "document subject", creator = "Me me me",
                       description = "this document is not empty",
                       created = time_now ) %>%
    print(target = filename )

  doc <- read_docx(path = filename)
  properties <- doc_properties(doc)

  expect_equal( properties$value[properties$tag %in% "title"], "title" )
  expect_equal( properties$value[properties$tag %in% "subject"], "document subject" )
  expect_equal( properties$value[properties$tag %in% "creator"], "Me me me" )
  expect_equal( properties$value[properties$tag %in% "description"], "this document is not empty" )
  expect_equal( properties$value[properties$tag %in% "created"], format( time_now, "%Y-%m-%dT%H:%M:%SZ") )

})



test_that("set pptx properties", {
  doc <- read_pptx()
  time_now <- Sys.time()
  filename <- tempfile(fileext = ".pptx")
  doc <- set_doc_properties(doc, title = "title",
                       subject = "document subject", creator = "Me me me",
                       description = "this document is not empty",
                       created = time_now )
  print(doc, target = filename )

  doc <- read_pptx(path = filename)
  properties <- doc_properties(doc)

  expect_equal( properties$value[properties$tag %in% "title"], "title" )
  expect_equal( properties$value[properties$tag %in% "subject"], "document subject" )
  expect_equal( properties$value[properties$tag %in% "creator"], "Me me me" )
  expect_equal( properties$value[properties$tag %in% "description"], "this document is not empty" )
  expect_equal( properties$value[properties$tag %in% "created"], format( time_now, "%Y-%m-%dT%H:%M:%SZ") )

})
