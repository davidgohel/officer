test_that("read docx properties", {
  doc <- read_docx(
    path = system.file(package = "officer", "doc_examples/example.docx")
  )
  properties <- doc_properties(doc)
  expect_equal(properties$value[properties$tag %in% "title"], "document title")
  expect_equal(
    properties$value[properties$tag %in% "subject"],
    "document subject"
  )
  expect_equal(properties$value[properties$tag %in% "creator"], "author")
  expect_equal(
    properties$value[properties$tag %in% "description"],
    "these are comments"
  )
  expect_equal(
    properties$value[properties$tag %in% "created"],
    "2017-04-26T13:10:00Z"
  )
})

test_that("read pptx properties", {
  doc <- read_pptx(
    path = system.file(package = "officer", "doc_examples/example.pptx")
  )
  properties <- doc_properties(doc)
  expect_equal(properties$value[properties$tag %in% "title"], "document title")
  expect_equal(
    properties$value[properties$tag %in% "subject"],
    "document subject"
  )
  expect_equal(properties$value[properties$tag %in% "creator"], "author")
  expect_equal(
    properties$value[properties$tag %in% "description"],
    "these are comments"
  )
  expect_equal(
    properties$value[properties$tag %in% "created"],
    "2017-04-27T11:29:40Z"
  )
})


test_that("set docx properties", {
  doc <- read_docx()
  time_now <- Sys.time()
  filename <- tempfile(fileext = ".docx")
  doc <- set_doc_properties(
    doc,
    title = "title",
    subject = "document subject",
    creator = "Me me me",
    description = "this document is not empty",
    created = time_now
  )
  dooc <- print(doc, target = filename)

  doc <- read_docx(path = filename)
  properties <- doc_properties(doc)

  expect_equal(properties$value[properties$tag %in% "title"], "title")
  expect_equal(
    properties$value[properties$tag %in% "subject"],
    "document subject"
  )
  expect_equal(properties$value[properties$tag %in% "creator"], "Me me me")
  expect_equal(
    properties$value[properties$tag %in% "description"],
    "this document is not empty"
  )
  expect_equal(
    properties$value[properties$tag %in% "created"],
    format(time_now, "%Y-%m-%dT%H:%M:%SZ", tz = "UTC")
  )
})

test_that("set custom properties", {
  filename <- tempfile(fileext = ".docx")
  doc <- read_docx()
  doc <- set_doc_properties(
    doc,
    coco = "coucou",
    zozo = "zuzu",
    isbn = NULL,
    bozo = NA_character_
  )
  print(doc, target = filename)

  doc <- read_docx(path = filename)
  properties <- doc_properties(doc)

  expect_equal(properties$value[properties$tag %in% "coco"], "coucou")
  expect_equal(properties$value[properties$tag %in% "zozo"], "zuzu")
  expect_equal(properties$value[properties$tag %in% "isbn"], "")
  expect_equal(properties$value[properties$tag %in% "bozo"], "")

  filename <- tempfile(fileext = ".pptx")
  doc <- read_pptx()
  doc <- set_doc_properties(doc, coco = "coucou", zozo = "zuzu")
  print(doc, target = filename)

  doc <- read_pptx(path = filename)
  properties <- doc_properties(doc)

  expect_equal(properties$value[properties$tag %in% "coco"], "coucou")
  expect_equal(properties$value[properties$tag %in% "zozo"], "zuzu")

  doc <- read_docx()
  filename <- tempfile(fileext = ".docx")
  doc <- set_doc_properties(doc, values = list(coco = "coucou", zozo = "zuzu"))
  print(doc, target = filename)

  doc <- read_docx(path = filename)
  properties <- doc_properties(doc)

  expect_equal(properties$value[properties$tag %in% "coco"], "coucou")
  expect_equal(properties$value[properties$tag %in% "zozo"], "zuzu")

  doc <- read_pptx()
  filename <- tempfile(fileext = ".pptx")
  doc <- set_doc_properties(doc, values = list(coco = "coucou", zozo = "zuzu"))
  print(doc, target = filename)

  doc <- read_pptx(path = filename)
  properties <- doc_properties(doc)

  expect_equal(properties$value[properties$tag %in% "coco"], "coucou")
  expect_equal(properties$value[properties$tag %in% "zozo"], "zuzu")
})


test_that("set pptx properties", {
  doc <- read_pptx()
  time_now <- Sys.time()
  filename <- tempfile(fileext = ".pptx")
  doc <- set_doc_properties(
    doc,
    title = "title",
    subject = "document subject",
    creator = "Me me me",
    description = "this document is not empty",
    created = time_now
  )
  print(doc, target = filename)

  doc <- read_pptx(path = filename)
  properties <- doc_properties(doc)

  expect_equal(properties$value[properties$tag %in% "title"], "title")
  expect_equal(
    properties$value[properties$tag %in% "subject"],
    "document subject"
  )
  expect_equal(properties$value[properties$tag %in% "creator"], "Me me me")
  expect_equal(
    properties$value[properties$tag %in% "description"],
    "this document is not empty"
  )
  expect_equal(
    properties$value[properties$tag %in% "created"],
    format(time_now, "%Y-%m-%dT%H:%M:%SZ", tz = "UTC")
  )
})

core_props_of <- function(path) {
  officer:::read_core_properties(unpack_folder(path, tempfile()))
}

new_docs <- function() {
  list(docx = read_docx(), pptx = read_pptx(), xlsx = read_xlsx())
}

parse_core_datetime <- function(x) {
  as.POSIXct(x, format = "%Y-%m-%dT%H:%M:%SZ", tz = "UTC")
}

test_that("created and modified are written in UTC", {
  withr::local_options(
    officer_datetime = as.POSIXct("2026-03-01 10:00:00", tz = "Europe/Paris")
  )
  docs <- new_docs()
  for (ext in names(docs)) {
    path <- tempfile(fileext = paste0(".", ext))
    print(docs[[ext]], target = path)
    props <- core_props_of(path)
    expect_equal(props["created", "value"], "2026-03-01T09:00:00Z", label = ext)
    expect_equal(props["modified", "value"], "2026-03-01T09:00:00Z", label = ext)
  }
})

test_that("created and modified are the current time in UTC by default", {
  withr::local_timezone("Europe/Paris")
  docs <- new_docs()
  for (ext in names(docs)) {
    path <- tempfile(fileext = paste0(".", ext))
    print(docs[[ext]], target = path)
    props <- core_props_of(path)
    for (prop in c("created", "modified")) {
      value <- parse_core_datetime(props[prop, "value"])
      delay <- abs(as.numeric(difftime(Sys.time(), value, units = "secs")))
      expect_lt(delay, 60, label = paste(ext, prop))
    }
  }
})

test_that("documents created from the templates have no creator", {
  docs <- new_docs()
  for (ext in names(docs)) {
    path <- tempfile(fileext = paste0(".", ext))
    print(docs[[ext]], target = path)
    expect_equal(core_props_of(path)["creator", "value"], "", label = ext)
  }
})

test_that("an existing document keeps its created property", {
  path <- tempfile(fileext = ".docx")
  doc <- read_docx(system.file(package = "officer", "doc_examples/example.docx"))
  print(doc, target = path)
  expect_equal(core_props_of(path)["created", "value"], "2017-04-26T13:10:00Z")
})

test_that("set_doc_properties() writes created in UTC", {
  created <- as.POSIXct("2026-03-01 10:00:00", tz = "Europe/Paris")
  doc <- set_doc_properties(read_docx(), created = created)
  path <- tempfile(fileext = ".docx")
  print(doc, target = path)
  expect_equal(core_props_of(path)["created", "value"], "2026-03-01T09:00:00Z")
})

test_that("option officer_datetime gives the same core properties", {
  withr::local_options(
    officer_datetime = as.POSIXct("2026-01-01", tz = "UTC")
  )
  core_xml <- function() {
    path <- tempfile(fileext = ".docx")
    print(read_docx(), target = path)
    readLines(file.path(unpack_folder(path, tempfile()), "docProps", "core.xml"))
  }
  expect_identical(core_xml(), core_xml())
})

test_that("lastModifiedBy is the name of the user", {
  path <- tempfile(fileext = ".docx")
  print(read_docx(), target = path)
  expect_equal(
    core_props_of(path)["lastModifiedBy", "value"],
    Sys.info()[["user"]]
  )
})

test_that("an invalid officer_datetime gives an error", {
  withr::local_options(officer_datetime = "2026-01-01")
  expect_error(read_docx(), "officer_datetime")
})

test_that("set xlsx properties", {
  doc <- read_xlsx()
  doc <- set_doc_properties(
    doc,
    title = "title",
    creator = "Me me me",
    coco = "coucou"
  )
  path <- tempfile(fileext = ".xlsx")
  print(doc, target = path)

  properties <- doc_properties(read_xlsx(path))
  expect_equal(properties$value[properties$tag %in% "title"], "title")
  expect_equal(properties$value[properties$tag %in% "creator"], "Me me me")
  expect_equal(properties$value[properties$tag %in% "coco"], "coucou")

  unpacked <- unpack_folder(path, tempfile())
  rels <- paste(readLines(file.path(unpacked, "_rels", ".rels"), warn = FALSE), collapse = "")
  expect_match(rels, "docProps/custom.xml", fixed = TRUE)
  ct <- paste(readLines(file.path(unpacked, "[Content_Types].xml"), warn = FALSE), collapse = "")
  expect_match(ct, "/docProps/custom.xml", fixed = TRUE)
})

test_that("templates carry no company and no last author", {
  docs <- new_docs()
  for (ext in names(docs)) {
    expect_equal(
      doc_properties(docs[[ext]])$value[doc_properties(docs[[ext]])$tag == "lastModifiedBy"],
      "",
      label = ext
    )
    path <- tempfile(fileext = paste0(".", ext))
    print(docs[[ext]], target = path)
    app <- xml2::read_xml(file.path(unpack_folder(path, tempfile()), "docProps", "app.xml"))
    company <- xml2::xml_find_first(app, "//d1:Company", ns = xml2::xml_ns(app))
    expect_equal(xml2::xml_text(company), "", label = ext)
  }
})

test_that("the same code produces the same file", {
  withr::local_options(
    officer_datetime = as.POSIXct("2026-01-01", tz = "UTC")
  )
  img <- file.path(R.home("doc"), "html", "logo.jpg")
  writers <- list(
    docx = function(path) {
      doc <- body_add_img(read_docx(), img, width = 1, height = 1)
      print(doc, target = path)
    },
    pptx = function(path) {
      doc <- add_slide(read_pptx(), "Title and Content", "Office Theme")
      doc <- ph_with(doc, external_img(img, width = 1, height = 1),
                     location = ph_location(left = 1, top = 1))
      print(doc, target = path)
    },
    xlsx = function(path) {
      doc <- sheet_write_data(read_xlsx(), head(iris), sheet = "Feuil1")
      print(doc, target = path)
    }
  )
  for (ext in names(writers)) {
    md5 <- vapply(c("UTC", "Pacific/Auckland"), function(tz) {
      withr::local_timezone(tz)
      path <- tempfile(fileext = paste0(".", ext))
      writers[[ext]](path)
      unname(tools::md5sum(path))
    }, character(1))
    expect_equal(md5[[1]], md5[[2]], label = ext)
  }
})

test_that("[Content_Types].xml is the first entry, whatever the locale", {
  old <- Sys.getlocale("LC_COLLATE")
  on.exit(Sys.setlocale("LC_COLLATE", old), add = TRUE)
  skip_if(
    identical(suppressWarnings(Sys.setlocale("LC_COLLATE", "fr_FR.UTF-8")), ""),
    "locale fr_FR.UTF-8 is not available"
  )
  path <- tempfile(fileext = ".docx")
  print(read_docx(), target = path)
  expect_equal(zip::zip_list(path)$filename[1], "[Content_Types].xml")
})
