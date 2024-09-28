
test_that("ph_labels with same order as in layout_properties()", {
  x <- read_pptx()
  l1 <- layout_rename_ph_labels(x, "Comparison")
  l2 <- layout_properties(x, "Comparison")$ph_label
  expect_equal(l1, l2)
})


test_that("incorrect inputs are detected", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()
  layout <- "Comparison"

  # unnamed args in renaming (dots)
  error_msg <- "Unnamed arguments are not allowed."
  expect_error(layout_rename_ph_labels(x, layout, NULL, "xxxx"), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, .dots = list("xxxx")), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, NULL, "xxx", .dots = list(a = "xxxx")), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, "xxx" = "a", .dots = list("xxxx")), error_msg)

  # unknown labels
  error_msg <- "Can't rename labels that don't exist."
  expect_error(layout_rename_ph_labels(x, layout, "xxxx" = "..."), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, "xxxx" = "...", "yyy" = "..."), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, .dots = list("xxxx" = "...")), error_msg)

  # unknown ids
  error_msg <- "Can't rename ids that don't exist."
  expect_error(layout_rename_ph_labels(x, layout, "1" = "..."), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, "1" = "...", "0" = "..."), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, .dots = list("1" = "...")), error_msg)

  # duplicate rename entries
  error_msg <- "Each id or label must only have one rename entry only."
  expect_error(layout_rename_ph_labels(x, layout, "Title 1" = "a", "Title 1" = "b"), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, "2" = "a", "2" = "b"), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, .dots = list("Title 1" = "a", "Title 1" = "b")), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, "2" = "a", .dots = list("2" = "b")), error_msg)

  # label and id collision
  error_msg <- "Either specify the label OR the id of the ph to rename, not both."
  expect_error(layout_rename_ph_labels(x, layout, "Title 1" = "a", "2" = "b"), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, "Title 1" = "a", "2" = "b"), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, .dots = list("Title 1" = "a", "2" = "b")), error_msg)
  expect_error(layout_rename_ph_labels(x, layout, "Date Placeholder 6" = "a", "7" = "b", .dots = list("Title 1" = "a", "2" = "b")), error_msg)
})


test_that("ph renaming works as expected", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()
  layout <- "Comparison"

  # rename using key-value pairs: 'old label' = 'new label' or 'id' = 'new label'
  layout_rename_ph_labels(x, layout, "Title 1" = "LABEL MATCHED") # label matching
  layout_rename_ph_labels(x, layout, "3" = "ID MATCHED") # id matching
  layout_rename_ph_labels(x, layout, "Date Placeholder 6" = "DATE", "8" = "FOOTER") # label and id
  layout_properties(x, layout)$ph_label

  x <- read_pptx()
  idx <- c(1, 2, 6, 7)
  l <- list("Date Placeholder 6" = "idx_6", "8" = "idx_7", "Title 1" = "idx_1", "3" = "idx_2") # as list
  layout_rename_ph_labels(x, layout, .dots = l)
  expect_equal(layout_properties(x, layout)$ph_label[idx], paste0("idx_", idx))

  x <- read_pptx()
  l <- c("Date Placeholder 6" = "idx_6", "8" = "idx_7", "Title 1" = "idx_1", "3" = "idx_2") # as vector
  layout_rename_ph_labels(x, layout, .dots = l)
  expect_equal(layout_properties(x, layout)$ph_label[idx], paste0("idx_", idx))

  x <- read_pptx()
  l <- list("Date Placeholder 6" = "idx_6", "3" = "idx_2")
  layout_rename_ph_labels(x, layout, "8" = "idx_7", "Title 1" = "idx_1", .dots = l) # mix ... and .dots
  expect_equal(layout_properties(x, layout)$ph_label[idx], paste0("idx_", idx))

  # rename via rhs assignment and ph index (not id!)
  x <- read_pptx()
  rhs <- LETTERS[1:8]
  layout_rename_ph_labels(x, layout) <- rhs
  expect_equal(layout_properties(x, layout)$ph_label, rhs)

  rhs <- paste("CHANGED", 1:3)
  ph_label_check <- layout_properties(x, layout)$ph_label
  ph_label_check[1:3] <- rhs
  layout_rename_ph_labels(x, layout)[1:3] <- rhs
  expect_equal(layout_properties(x, layout)$ph_label, ph_label_check)

  # rename via rhs assignment and ph id (not index)
  lp_old <- ph_label_check <- layout_properties(x, layout)
  ids <- c(2, 4, 5)
  idx <- match(ids, lp_old$id) # row in layout properties
  rhs <- paste("ID =", ids)
  ph_label_check <- lp_old$ph_label
  ph_label_check[idx] <- rhs
  layout_rename_ph_labels(x, layout, id = ids) <- rhs
  lp_new <- layout_properties(x, layout)
  expect_equal(lp_new$ph_label, ph_label_check)
})


test_that("renaming duplicate labels replaces 1st occurrence only", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  file <- test_path("docs_dir", "test-pptx-dedupe-ph.pptx")
  x <- read_pptx(file)

  # rename first label occurrence only and issue warning (1 duped label)
  layout <- "2-dupes"
  ph_label_check <- layout_properties(x, layout)$ph_label
  idx <- which(ph_label_check == "Content 7") # exists twice
  new_value <- "xxxx"
  warn_msg <- "When renaming a label with duplicates, only the first occurrence is renamed."
  expect_warning(layout_rename_ph_labels(x, layout, "Content 7" = new_value), warn_msg)
  ph_label_new <- layout_properties(x, layout)$ph_label
  ph_label_check[idx[1]] <- new_value # only first occurrence is replaced
  expect_equal(ph_label_check, ph_label_new)

  # rename first label occurrence only and issue warning (2 duped labels)
  layout <- "2x2-dupes"
  ph_label_check <- layout_properties(x, layout)$ph_label
  ii <- which(duplicated(ph_label_check, fromLast = TRUE)) # index of 1st occurrence
  dupes <- ph_label_check[ii]
  vals <- LETTERS[seq_along(dupes)]
  names(vals) <- dupes
  warn_msg_1 <- "When renaming a label with duplicates, only the first occurrence is renamed."
  warn_msg_2 <- "Renaming 2 ph labels with duplicates"
  warn_msg <- paste0(warn_msg_1, ".*", warn_msg_2)
  expect_warning(layout_rename_ph_labels(x, layout, .dots = vals), warn_msg)
  ph_label_check[ii] <- vals
  ph_label_new <- layout_properties(x, layout)$ph_label
  expect_equal(ph_label_check, ph_label_new)
})

