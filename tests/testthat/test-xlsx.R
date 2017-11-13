context("xlsx document")


test_that("create and manipulate sheet", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "sheet1")
  doc <- sheet_select(doc, sheet = "sheet1")

  expect_true( "sheet1" %in% doc$worksheets$sheet_names() )

  xml_sheets <- list.files( file.path(doc$package_dir, "xl", "worksheets"),
                            pattern = "\\.xml$" )
  expect_equal( length(doc), 2 )
  expect_equal( xml_sheets, c("sheet1.xml", "sheet2.xml") )

  sheet_id <- doc$worksheets$get_sheet_id("sheet1")
  wb_view <- xml_find_first(doc$worksheets$get(), "d1:bookViews/d1:workbookView")
  expect_equal(as.integer(xml_attr(wb_view, "activeTab")), sheet_id - 1)

})



test_that("add data.frame to sheet", {
  doc <- read_xlsx()
  doc <- add_sheet(doc, label = "sheet1")
  doc <- sheet_select(doc, sheet = "sheet1")
  doc <- sheet_add_table(doc, sheet = "sheet1", iris, at_row = 1, at_col = 1)

  xml_sheet <- doc$sheets$get_sheet(2)$get()
  row_data <- xml_find_all(xml_sheet, "d1:sheetData/d1:row")

  expect_equal( xml_text( xml_children(row_data[[1]]) ), names(iris))
  expect_equal( xml_find_all(xml_sheet, "d1:sheetData/d1:row") %>% length(), 151)
})






