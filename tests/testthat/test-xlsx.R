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





