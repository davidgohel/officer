test_that("create and manipulate sheet", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "sheet1")
  doc <- sheet_select(doc, sheet = "sheet1")

  expect_true("sheet1" %in% doc$worksheets$sheet_names())

  xml_sheets <- list.files(
    file.path(doc$package_dir, "xl", "worksheets"),
    pattern = "\\.xml$"
  )
  expect_equal(length(doc), 2)
  expect_equal(xml_sheets, c("sheet1.xml", "sheet2.xml"))

  sheet_id <- doc$worksheets$get_sheet_id("sheet1")
  wb_view <- xml_find_first(
    doc$worksheets$get(),
    "d1:bookViews/d1:workbookView"
  )
  expect_equal(as.integer(xml_attr(wb_view, "activeTab")), sheet_id - 1)
})

test_that("sheet_select deselects other sheets", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "new")
  doc <- sheet_select(doc, sheet = "new")
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  sheet1_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet1.xml")
  )
  sv1 <- xml_find_first(sheet1_xml, "d1:sheetViews/d1:sheetView", ns = ns)
  expect_equal(xml_attr(sv1, "tabSelected"), "0")
})

test_that("sheet_write_data writes correct cells", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "data")
  doc <- sheet_write_data(doc, value = head(iris, 3), sheet = "data")
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  sheet_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet2.xml")
  )
  rows <- xml_find_all(sheet_xml, "d1:sheetData/d1:row", ns = ns)

  # header + 3 data rows
  expect_equal(length(rows), 4L)

  # header row has 5 columns
  header_cells <- xml_find_all(rows[[1]], "d1:c", ns = ns)
  expect_equal(length(header_cells), 5L)

  # first header is Sepal.Length
  first_val <- xml_text(xml_find_first(
    header_cells[[1]],
    "d1:is/d1:t",
    ns = ns
  ))
  expect_equal(first_val, "Sepal.Length")

  # numeric value in B2
  b2 <- xml_find_first(rows[[2]], "d1:c[@r='B2']/d1:v", ns = ns)
  expect_equal(xml_text(b2), "3.5")

  # text value in E2 (Species)
  e2 <- xml_find_first(rows[[2]], "d1:c[@r='E2']/d1:is/d1:t", ns = ns)
  expect_equal(xml_text(e2), "setosa")
})

test_that("sheet_write_data with start_row and start_col", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "data")
  doc <- sheet_write_data(
    doc,
    value = data.frame(a = 1:2, b = 3:4),
    sheet = "data",
    start_row = 5,
    start_col = 3
  )
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  sheet_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet2.xml")
  )

  # header at row 5, col C
  c5 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='5']/d1:c[@r='C5']",
    ns = ns
  )
  expect_false(inherits(c5, "xml_missing"))
  expect_equal(
    xml_text(xml_find_first(c5, "d1:is/d1:t", ns = ns)),
    "a"
  )

  # data at D6
  d6 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='6']/d1:c[@r='D6']/d1:v",
    ns = ns
  )
  expect_equal(xml_text(d6), "3")
})

test_that("sheet_write_data merges two datasets", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "data")
  doc <- sheet_write_data(
    doc,
    value = data.frame(x = 1:2),
    sheet = "data"
  )
  doc <- sheet_write_data(
    doc,
    value = data.frame(y = 10:11),
    sheet = "data",
    start_col = 3
  )
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  sheet_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet2.xml")
  )

  # row 1 has both headers
  row1_cells <- xml_find_all(
    sheet_xml,
    "d1:sheetData/d1:row[@r='1']/d1:c",
    ns = ns
  )
  expect_equal(length(row1_cells), 2L)

  # A2 has 1, C2 has 10
  a2 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='2']/d1:c[@r='A2']/d1:v",
    ns = ns
  )
  c2 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='2']/d1:c[@r='C2']/d1:v",
    ns = ns
  )
  expect_equal(xml_text(a2), "1")
  expect_equal(xml_text(c2), "10")
})

test_that("sheet_write_data handles NA", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "data")
  doc <- sheet_write_data(
    doc,
    value = data.frame(a = c(1, NA), b = c("x", NA)),
    sheet = "data"
  )
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  sheet_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet2.xml")
  )

  # A3 is empty (NA numeric)
  a3 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='3']/d1:c[@r='A3']",
    ns = ns
  )
  expect_false(inherits(a3, "xml_missing"))
  a3_val <- xml_find_first(a3, "d1:v", ns = ns)
  expect_true(inherits(a3_val, "xml_missing"))
})

test_that("sheet_write_data handles Date columns", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "data")
  df <- data.frame(
    d = as.Date(c("2024-01-15", "2024-06-30", NA)),
    x = 1:3
  )
  doc <- sheet_write_data(doc, value = df, sheet = "data")
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  # styles.xml has numFmtId 14 (date)
  styles_xml <- paste(
    readLines(
      file.path(unpack_dir, "xl/styles.xml"),
      warn = FALSE
    ),
    collapse = ""
  )
  expect_true(grepl("numFmtId=\"14\"", styles_xml))

  # date cells have s= attribute
  sheet_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet2.xml")
  )
  a2 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='2']/d1:c[@r='A2']",
    ns = ns
  )
  expect_false(is.na(xml_attr(a2, "s")))

  # date value is Excel serial number (2024-01-15 = 45306)
  a2_val <- as.numeric(xml_text(xml_find_first(a2, "d1:v", ns = ns)))
  expect_equal(a2_val, as.numeric(as.Date("2024-01-15")) + 25569)

  # NA date produces empty cell
  a4 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='4']/d1:c[@r='A4']",
    ns = ns
  )
  a4_val <- xml_find_first(a4, "d1:v", ns = ns)
  expect_true(inherits(a4_val, "xml_missing"))
})

test_that("sheet_write_data handles POSIXct columns", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "data")
  df <- data.frame(
    dt = as.POSIXct(
      c("2024-01-15 10:30:00", "2024-06-30 14:00:00", NA),
      tz = "UTC"
    ),
    x = 1:3
  )
  doc <- sheet_write_data(doc, value = df, sheet = "data")
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  # styles.xml has numFmtId 22 (datetime)
  styles_xml <- paste(
    readLines(
      file.path(unpack_dir, "xl/styles.xml"),
      warn = FALSE
    ),
    collapse = ""
  )
  expect_true(grepl("numFmtId=\"22\"", styles_xml))

  # datetime cells have s= attribute
  sheet_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet2.xml")
  )
  a2 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='2']/d1:c[@r='A2']",
    ns = ns
  )
  expect_false(is.na(xml_attr(a2, "s")))

  # value is a serial number > 45306 (date part) with fractional time
  a2_val <- as.numeric(xml_text(xml_find_first(a2, "d1:v", ns = ns)))
  expect_gt(a2_val, 45306)
  expect_true(a2_val != round(a2_val)) # has fractional part (time)

  # NA datetime produces empty cell
  a4 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='4']/d1:c[@r='A4']",
    ns = ns
  )
  a4_val <- xml_find_first(a4, "d1:v", ns = ns)
  expect_true(inherits(a4_val, "xml_missing"))
})

test_that("sheet_write_data handles logical columns", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "data")
  df <- data.frame(flag = c(TRUE, FALSE, NA))
  doc <- sheet_write_data(doc, value = df, sheet = "data")
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  sheet_xml <- read_xml(
    file.path(unpack_dir, "xl/worksheets/sheet2.xml")
  )

  # TRUE -> t="b", v=1
  a2 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='2']/d1:c[@r='A2']",
    ns = ns
  )
  expect_equal(xml_attr(a2, "t"), "b")
  expect_equal(xml_text(xml_find_first(a2, "d1:v", ns = ns)), "1")

  # FALSE -> t="b", v=0
  a3 <- xml_find_first(
    sheet_xml,
    "d1:sheetData/d1:row[@r='3']/d1:c[@r='A3']",
    ns = ns
  )
  expect_equal(xml_text(xml_find_first(a3, "d1:v", ns = ns)), "0")
})

test_that("sheet_add_drawing creates drawing infrastructure", {
  skip_if_not_installed("mschart", minimum_version = "0.4.2")
  library(mschart)

  my_chart <- ms_barchart(
    data = data.frame(
      x = c("A", "B"),
      y = c(1, 2),
      group = rep("s1", 2)
    ),
    x = "x",
    y = "y",
    group = "group"
  )

  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "chart")
  doc <- sheet_add_drawing(
    doc,
    value = my_chart,
    sheet = "chart"
  )
  out <- print(doc, target = tempfile(fileext = ".xlsx"))

  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  # drawing file exists
  expect_true(
    file.exists(file.path(unpack_dir, "xl/drawings/drawing1.xml"))
  )

  # chart file exists
  chart_files <- list.files(
    file.path(unpack_dir, "xl/charts"),
    pattern = "^chart.*\\.xml$"
  )
  expect_gte(length(chart_files), 1L)

  # sheet has drawing ref before extLst
  sheet_str <- paste(
    readLines(
      file.path(unpack_dir, "xl/worksheets/sheet2.xml"),
      warn = FALSE
    ),
    collapse = ""
  )
  drawing_pos <- regexpr("<drawing", sheet_str)
  extlst_pos <- regexpr("<extLst", sheet_str)
  if (extlst_pos > 0) {
    expect_true(drawing_pos < extlst_pos)
  }

  # drawing rels points to chart
  drawing_rels <- paste(
    readLines(
      file.path(unpack_dir, "xl/drawings/_rels/drawing1.xml.rels"),
      warn = FALSE
    ),
    collapse = ""
  )
  expect_true(grepl("chart", drawing_rels))

  # no externalData in chart
  chart_str <- paste(
    readLines(
      file.path(unpack_dir, "xl/charts", chart_files[1]),
      warn = FALSE
    ),
    collapse = ""
  )
  expect_false(grepl("externalData", chart_str))
})

sheet_cells_xml <- function(doc, sheet_xml_name = "sheet2.xml") {
  out <- print(doc, target = tempfile(fileext = ".xlsx"))
  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)
  paste(readLines(file.path(unpack_dir, "xl/worksheets", sheet_xml_name)),
        collapse = "\n")
}

test_that("sheet_write_data dispatches on character (vertical)", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "chr")
  doc <- sheet_write_data(doc, value = c("A", "B", "C"),
                          sheet = "chr", start_row = 2, start_col = 3)
  xml <- sheet_cells_xml(doc)
  expect_match(xml, "<c r=\"C2\" t=\"inlineStr\"><is><t[^>]*>A</t>", fixed = FALSE)
  expect_match(xml, "<c r=\"C3\" t=\"inlineStr\"><is><t[^>]*>B</t>", fixed = FALSE)
  expect_match(xml, "<c r=\"C4\" t=\"inlineStr\"><is><t[^>]*>C</t>", fixed = FALSE)
})

test_that("sheet_write_data dispatches on character (horizontal)", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "chr")
  doc <- sheet_write_data(doc, value = c("X", "Y", "Z"),
                          sheet = "chr", start_row = 1, start_col = 1,
                          direction = "horizontal")
  xml <- sheet_cells_xml(doc)
  expect_match(xml, "r=\"A1\"[^>]*>[^<]*<is><t[^>]*>X</t>", fixed = FALSE)
  expect_match(xml, "r=\"B1\"[^>]*>[^<]*<is><t[^>]*>Y</t>", fixed = FALSE)
  expect_match(xml, "r=\"C1\"[^>]*>[^<]*<is><t[^>]*>Z</t>", fixed = FALSE)
})

test_that("sheet_write_data on fpar emits richtext runs", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "fp")
  f <- fpar(
    ftext("bold ", fp_text(bold = TRUE, color = "red")),
    ftext("italic ", fp_text(italic = TRUE)),
    ftext("H", fp_text()),
    ftext("2", fp_text(vertical.align = "subscript")),
    ftext("O", fp_text())
  )
  doc <- sheet_write_data(doc, value = f, sheet = "fp",
                          start_row = 1, start_col = 1)
  xml <- sheet_cells_xml(doc)
  expect_match(xml, "<c r=\"A1\" t=\"inlineStr\">", fixed = TRUE)
  expect_match(xml, "<b/>", fixed = TRUE)
  expect_match(xml, "rgb=\"FFFF0000\"", fixed = TRUE)
  expect_match(xml, "<i/>", fixed = TRUE)
  expect_match(xml, "<vertAlign val=\"subscript\"/>", fixed = TRUE)
})

test_that("sheet_write_data on block_list stacks fpars vertically", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "bl")
  bl <- block_list(
    fpar(ftext("line 1", fp_text(bold = TRUE))),
    fpar(ftext("line 2", fp_text(italic = TRUE)))
  )
  doc <- sheet_write_data(doc, value = bl, sheet = "bl",
                          start_row = 5, start_col = 2)
  xml <- sheet_cells_xml(doc)
  expect_match(xml, "r=\"B5\"[^>]*t=\"inlineStr\"", fixed = FALSE)
  expect_match(xml, "r=\"B6\"[^>]*t=\"inlineStr\"", fixed = FALSE)
  expect_match(xml, "line 1", fixed = TRUE)
  expect_match(xml, "line 2", fixed = TRUE)
})

test_that("sheet_write_data default method errors on unsupported input", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "x")
  expect_error(sheet_write_data(doc, value = 42, sheet = "x"),
               "method")
})

test_that("sheet_add_drawing(.external_img) embeds images", {
  img <- file.path(R.home("doc"), "html", "logo.jpg")
  skip_if_not(file.exists(img), message = "no sample image available")

  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "pics")
  doc <- sheet_add_drawing(doc, sheet = "pics",
                           value = external_img(img, width = 2, height = 1.5),
                           left = 2, top = 1)
  doc <- sheet_add_drawing(doc, sheet = "pics",
                           value = external_img(img, width = 3, height = 2),
                           left = 2, top = 4)

  out <- print(doc, target = tempfile(fileext = ".xlsx"))
  unpack_dir <- tempfile()
  unpack_folder(out, unpack_dir)

  media <- list.files(file.path(unpack_dir, "xl/media"))
  expect_length(media, 2)

  drawings <- list.files(file.path(unpack_dir, "xl/drawings"),
                         pattern = "\\.xml$")
  expect_length(drawings, 1)

  drawing_xml <- paste(readLines(file.path(unpack_dir, "xl/drawings",
                                           drawings[1])),
                       collapse = "\n")
  # both anchors in the single drawing
  anchors <- regmatches(drawing_xml,
                        gregexpr("<xdr:absoluteAnchor",
                                 drawing_xml, perl = TRUE))[[1]]
  expect_length(anchors, 2)
})

