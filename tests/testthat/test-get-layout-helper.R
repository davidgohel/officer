

# test_that("incorrect inputs are detected", {
#   opts <- options(cli.num_colors = 1) # suppress colors for error message check
#   on.exit(options(opts))
#
#   x <- read_pptx()
#   layout <- "Comparison"
#
#   get_layout("xxx", "xxx")
#
#   # label and id collision
#   error_msg <- "Either specify the label OR the id of the ph to rename, not both."
#   expect_error(layout_rename_ph_labels(x, layout, "Title 1" = "a", "2" = "b"), error_msg)
#   expect_error(layout_rename_ph_labels(x, layout, "Title 1" = "a", "2" = "b"), error_msg)
#   expect_error(layout_rename_ph_labels(x, layout, .dots = list("Title 1" = "a", "2" = "b")), error_msg)
#   expect_error(layout_rename_ph_labels(x, layout, "Date Placeholder 6" = "a", "7" = "b", .dots = list("Title 1" = "a", "2" = "b")), error_msg)
# })


