test_that("console printing", {
  x <- read_xlsx()
  expect_output(print(x), regexp = "^xlsx document with 1 sheet")
})

test_that("check extention and print document", {
  x <- read_xlsx()
  print(x, target = "print.xlsx")
  expect_true(file.exists("print.xlsx"))
  expect_error(print(x, target = "print.xlsxxxx"))
})


unlink("*.xlsx")

test_that("ooxml_chart_uris returns paired URIs for both kinds", {
  dml <- ooxml_chart_uris("drawingml")
  expect_named(dml, c("rel_type", "graphic_uri"))
  expect_match(dml$rel_type, "/officeDocument/2006/relationships/chart$")
  expect_match(dml$graphic_uri, "/drawingml/2006/chart$")

  cx <- ooxml_chart_uris("chartex")
  expect_match(cx$rel_type, "/2014/relationships/chartEx$")
  expect_match(cx$graphic_uri, "/drawing/2014/chartex$")

  expect_identical(ooxml_chart_uris(), dml)
  expect_error(ooxml_chart_uris("nope"))
})

test_that("default_chart_inner dispatches on graphic_uri", {
  u <- ooxml_chart_uris("drawingml")
  dml <- officer:::default_chart_inner(u$graphic_uri, "rId7")
  expect_match(dml, "^<c:chart ")
  expect_match(dml, "xmlns:c=\"[^\"]+/drawingml/2006/chart\"")
  expect_match(dml, "r:id=\"rId7\"")

  v <- ooxml_chart_uris("chartex")
  cx <- officer:::default_chart_inner(v$graphic_uri, "rId7")
  expect_match(cx, "^<cx:chart ")
  expect_match(cx, "xmlns:cx=\"[^\"]+/drawing/2014/chartex\"")
  expect_match(cx, "r:id=\"rId7\"")
})

test_that("parse_cell_ref handles A1-style refs", {
  expect_equal(officer:::parse_cell_ref("A1"), list(col = 0L, row = 0L))
  expect_equal(officer:::parse_cell_ref("C4"), list(col = 2L, row = 3L))
  expect_equal(officer:::parse_cell_ref("aa12"), list(col = 26L, row = 11L))
  expect_equal(officer:::parse_cell_ref("Z99"), list(col = 25L, row = 98L))
  expect_error(officer:::parse_cell_ref("3A"), "invalid")
  expect_error(officer:::parse_cell_ref(""), "non-empty")
  expect_error(officer:::parse_cell_ref(c("A1", "B2")), "single")
})

# Helper: get a fresh xlsx_drawing on sheet 1 of a freshly-written xlsx,
# and return the inspector that reads the drawing XML back.
fresh_drawing <- function() {
  tmp <- tempfile(fileext = ".xlsx")
  x <- read_xlsx()
  print(x, target = tmp)
  x2 <- read_xlsx(tmp)
  sheet_obj <- x2$sheets$get_sheet(1)
  drw <- xlsx_drawing$new(x2$package_dir, sheet_obj, x2$content_type)
  list(drw = drw, doc = function() drw$get())
}

test_that("add_chart_anchor emits twoCellAnchor with from + to", {
  fd <- fresh_drawing()
  rid <- fd$drw$add_chart_rel("chart1.xml")
  fd$drw$add_chart_anchor(rid, from = "B2", to = "H20")
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  anchor <- xml2::xml_find_first(fd$doc(), "//xdr:twoCellAnchor", ns)
  expect_false(inherits(anchor, "xml_missing"))
  expect_equal(xml2::xml_attr(anchor, "editAs"), "twoCell")
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:from/xdr:col", ns)),
    "1"
  )
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:from/xdr:row", ns)),
    "1"
  )
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:to/xdr:col", ns)),
    "7"
  )
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:to/xdr:row", ns)),
    "19"
  )
})

test_that("add_chart_anchor honours edit_as", {
  fd <- fresh_drawing()
  rid <- fd$drw$add_chart_rel("chart1.xml")
  fd$drw$add_chart_anchor(rid, from = "A1", to = "B2", edit_as = "absolute")
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  expect_equal(
    xml2::xml_attr(
      xml2::xml_find_first(fd$doc(), "//xdr:twoCellAnchor", ns),
      "editAs"
    ),
    "absolute"
  )
})

test_that("add_chart_anchor emits oneCellAnchor with from only", {
  fd <- fresh_drawing()
  rid <- fd$drw$add_chart_rel("chart1.xml")
  fd$drw$add_chart_anchor(rid, from = "C4", width = 5, height = 3)
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  anchor <- xml2::xml_find_first(fd$doc(), "//xdr:oneCellAnchor", ns)
  expect_false(inherits(anchor, "xml_missing"))
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:from/xdr:col", ns)),
    "2"
  )
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:from/xdr:row", ns)),
    "3"
  )
  ext <- xml2::xml_find_first(anchor, "xdr:ext", ns)
  expect_equal(xml2::xml_attr(ext, "cx"), as.character(5 * 914400))
  expect_equal(xml2::xml_attr(ext, "cy"), as.character(3 * 914400))
})

test_that("add_chart_anchor keeps absoluteAnchor when from/to omitted", {
  fd <- fresh_drawing()
  rid <- fd$drw$add_chart_rel("chart1.xml")
  fd$drw$add_chart_anchor(rid, left = 0.5, top = 0.5, width = 4, height = 3)
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  expect_false(inherits(
    xml2::xml_find_first(fd$doc(), "//xdr:absoluteAnchor", ns),
    "xml_missing"
  ))
})

test_that("add_chart_anchor errors on `to` without `from`", {
  fd <- fresh_drawing()
  rid <- fd$drw$add_chart_rel("chart1.xml")
  expect_error(
    fd$drw$add_chart_anchor(rid, to = "B2"),
    "`to` requires `from`"
  )
})

test_that("add_image_anchor emits twoCellAnchor with from + to", {
  fd <- fresh_drawing()
  rid <- fd$drw$add_image_rel("image1.png")
  fd$drw$add_image_anchor(rid, from = "B2", to = "F10")
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  anchor <- xml2::xml_find_first(fd$doc(), "//xdr:twoCellAnchor", ns)
  expect_false(inherits(anchor, "xml_missing"))
  pic <- xml2::xml_find_first(anchor, "xdr:pic", ns)
  expect_false(inherits(pic, "xml_missing"))
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:to/xdr:col", ns)),
    "5"
  )
})

test_that("add_image_anchor emits oneCellAnchor with from only", {
  fd <- fresh_drawing()
  rid <- fd$drw$add_image_rel("image1.png")
  fd$drw$add_image_anchor(rid, from = "C4", width = 2, height = 1.5)
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  expect_false(inherits(
    xml2::xml_find_first(fd$doc(), "//xdr:oneCellAnchor", ns),
    "xml_missing"
  ))
})

test_that("parse_anchor decodes user-facing anchor strings", {
  expect_equal(officer:::parse_anchor(NULL), list(from = NULL, to = NULL))
  expect_equal(officer:::parse_anchor("C4"), list(from = "C4", to = NULL))
  expect_equal(
    officer:::parse_anchor("B2:H20"),
    list(from = "B2", to = "H20")
  )
  expect_error(officer:::parse_anchor(42), "single string")
  expect_error(officer:::parse_anchor("B2:"), "FROM:TO")
  expect_error(officer:::parse_anchor(""), "FROM:TO|single string")
})

test_that("sheet_add_drawing(external_img, anchor = range) uses twoCellAnchor", {
  img <- system.file(
    "medias",
    "figures",
    "body_add_doc_1.png",
    package = "officer"
  )
  skip_if_not(nzchar(img) && file.exists(img))
  x <- read_xlsx()
  x <- add_sheet(x, label = "pics")
  x <- sheet_add_drawing(
    x,
    sheet = "pics",
    value = external_img(img, width = 2, height = 2),
    anchor = "B2:F10"
  )
  out <- tempfile(fileext = ".xlsx")
  print(x, target = out)
  pkg <- tempfile()
  dir.create(pkg)
  zip::unzip(out, exdir = pkg)
  drw <- list.files(
    file.path(pkg, "xl", "drawings"),
    pattern = "\\.xml$",
    full.names = TRUE
  )[1]
  doc <- xml2::read_xml(drw)
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  anchor <- xml2::xml_find_first(doc, "//xdr:twoCellAnchor", ns)
  expect_false(inherits(anchor, "xml_missing"))
  expect_equal(
    xml2::xml_text(xml2::xml_find_first(anchor, "xdr:to/xdr:col", ns)),
    "5"
  )
})

test_that("sheet_add_drawing(external_img, anchor = single cell) uses oneCellAnchor", {
  img <- system.file(
    "medias",
    "figures",
    "body_add_doc_1.png",
    package = "officer"
  )
  skip_if_not(nzchar(img) && file.exists(img))
  x <- read_xlsx()
  x <- add_sheet(x, label = "pics")
  x <- sheet_add_drawing(
    x,
    sheet = "pics",
    value = external_img(img, width = 2, height = 2),
    anchor = "C4"
  )
  out <- tempfile(fileext = ".xlsx")
  print(x, target = out)
  pkg <- tempfile()
  dir.create(pkg)
  zip::unzip(out, exdir = pkg)
  drw <- list.files(
    file.path(pkg, "xl", "drawings"),
    pattern = "\\.xml$",
    full.names = TRUE
  )[1]
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  expect_false(inherits(
    xml2::xml_find_first(xml2::read_xml(drw), "//xdr:oneCellAnchor", ns),
    "xml_missing"
  ))
})

test_that("sheet_add_drawing(external_img) without anchor stays absolute", {
  img <- system.file(
    "medias",
    "figures",
    "body_add_doc_1.png",
    package = "officer"
  )
  skip_if_not(nzchar(img) && file.exists(img))
  x <- read_xlsx()
  x <- add_sheet(x, label = "pics")
  x <- sheet_add_drawing(
    x,
    sheet = "pics",
    value = external_img(img, width = 2, height = 2)
  )
  out <- tempfile(fileext = ".xlsx")
  print(x, target = out)
  pkg <- tempfile()
  dir.create(pkg)
  zip::unzip(out, exdir = pkg)
  drw <- list.files(
    file.path(pkg, "xl", "drawings"),
    pattern = "\\.xml$",
    full.names = TRUE
  )[1]
  ns <- c(
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  )
  expect_false(inherits(
    xml2::xml_find_first(xml2::read_xml(drw), "//xdr:absoluteAnchor", ns),
    "xml_missing"
  ))
})
