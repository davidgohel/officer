test_that("incorrect input formats are detected", {
  expect_error(
    layout_dedupe_ph_labels("file.xxx"),
    regexp = "'x' must be an 'rpptx' object"
  )
  expect_error(
    has_ph_dupes("file.xxx"),
    regexp = "'x' must be an 'rpptx' object"
  )
  expect_error(
    .dedupe_phs_in_layout("file.xxx"),
    regexp = "'layout_file' must be an .xml file"
  )
})


test_that("handling ph dupes function works when there are none", {
  x <- read_pptx() # sample PPTX has no dupes
  expect_false(has_ph_dupes(x))
  . <- capture.output(expect_no_error({
    layout_dedupe_ph_labels(x, print_info = TRUE)
    layout_dedupe_ph_labels(x, action = "rename", print_info = TRUE)
    layout_dedupe_ph_labels(x, action = "delete", print_info = TRUE)
  }))
})


test_that("handling ph dupes works", {
  file_in <- test_path("docs_dir/test-pptx-dedupe-ph.pptx")

  # referencing a duplicate placeholder via its label should throw an error.
  # if this should change for some reason, the test fails as we would need to
  # check if deduplication is still relevant
  x <- read_pptx(file_in)
  x <- add_slide(x, layout = "2x2-dupes", master = "Master1")
  expect_no_error(ph_with(x, "abc", ph_location_label(ph_label = "Title 1")))
  expect_error(ph_with(x, "abc", ph_location_label(ph_label = "Content")))

  # action = detect
  x_det <- read_pptx(file_in)
  expect_true(has_ph_dupes(x_det))
  out <- capture.output({
    x_det <- layout_dedupe_ph_labels(x_det)
  })
  expect_true(any(grepl("Content 7.1", out)))
  expect_true(has_ph_dupes(x_det))

  # action = rename
  x_rename <- read_pptx(file_in)
  before <- x_rename$slideLayouts$get_xfrm_data()$ph_label
  out <- capture.output({
    x_rename <- layout_dedupe_ph_labels(
      x_rename,
      action = "rename",
      print_info = TRUE
    )
  })
  expect_true(any(grepl("Content 7", out)))
  expect_true(any(grepl("Content 7.1", out)))
  after <- x_rename$slideLayouts$get_xfrm_data()$ph_label
  expect_false(has_ph_dupes(x_rename))
  expect_true(any(before != after))
  expect_equal(length(before), length(after))

  # action = delete
  x_delete <- read_pptx(file_in)
  before <- x_delete$slideLayouts$get_xfrm_data()$ph_label
  out <- capture.output({
    x_delete <- layout_dedupe_ph_labels(
      x_delete,
      action = "delete",
      print_info = TRUE
    )
  })
  expect_true(any(grepl("Content 7", out)))
  after <- x_delete$slideLayouts$get_xfrm_data()$ph_label
  expect_false(has_ph_dupes(x_delete))
  expect_gt(length(before), length(after))

  # annotate base should work with deduped phs
  file <- tempfile(fileext = ".pptx")
  output_file <- tempfile(fileext = ".pptx")
  print(x_rename, target = file)
  expect_no_error(annotate_base(file, output_file = output_file))
  print(x_delete, target = file)
  expect_no_error(annotate_base(file, output_file = output_file))
})
