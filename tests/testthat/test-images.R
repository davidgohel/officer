img.file <- file.path(R.home(component = "doc"), "html", "logo.jpg")
svg_file <- file.path(R.home(component = "doc"), "html", "Rlogo.svg")

ext_img <- external_img(img.file)
ext_svg <- external_img(svg_file)

test_that("add image in HTML", {
  expect_match(
    to_html(ext_svg),
    "<img style=\"vertical-align:middle;width:36px;height:14px;\" src=\"data:image/svg\\+xml;base64,"
  )
  expect_match(
    to_html(ext_img),
    "<img style=\"vertical-align:middle;width:36px;height:14px;\" src=\"data:image/jpeg;base64,"
  )
})

test_that("add image in docx", {
  x <- read_docx()
  x <- body_add_fpar(x, fpar(ext_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  rel_df <- x$doc_obj$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)image$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 1)
  if (nrow(subset_rel) > 0) {
    body <- docx_body_xml(x)
    node_blip <- xml_find_first(body, "//a:blip")
    expect_false(inherits(node_blip, "xml_missing"))
    expect_true(all(xml_attr(node_blip, "embed") %in% subset_rel$id))
  }
})

pic <- file.path(R.home("doc"), "html", "logo.jpg" )
base_dir <- tempfile()
file1 <- file.path(base_dir, "dir1", "logo1.jpg")
file2 <- file.path(base_dir, "dir2", "logo1.jpg")
file3 <- file.path(base_dir, "dir2", "logo2.jpg")
dir.create(file.path(base_dir, "dir1"), recursive = TRUE)
dir.create(file.path(base_dir, "dir2"), recursive = TRUE)
file.copy(pic, file1)
file.copy(pic, file2)
file.copy(pic, file3)

test_that("add multiple images in docx", {
  x <- read_docx()
  x <- body_add_img(x, src = file1, width = 1, height = 1)
  x <- body_add_img(x, src = file2, width = 1, height = 1)
  x <- body_add_img(x, src = file3, width = 1, height = 1)
  docx_file <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = docx_file)
  rel_df <- x$doc_obj$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)image$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 1)
  if (nrow(subset_rel) > 0) {
    body <- docx_body_xml(x)
    node_blip <- xml_find_all(body, "//a:blip")
    expect_length(node_blip, 3)
    expect_true(all(xml_attr(node_blip, "embed") %in% subset_rel$id))
  }
})

test_that("add image in pptx", {
  x <- read_pptx()
  x <- add_slide(x)
  x <- ph_with(x, ext_img, location = ph_location_type())
  filename <- print(x, target = tempfile(fileext = ".pptx"))

  x <- read_pptx(path = filename)
  slide <- x$slide$get_slide(x$cursor)

  rel_df <- slide$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)image$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 1)
  if (nrow(subset_rel) > 0) {

    node_blip <- xml_find_first(slide$get(), "//a:blip")
    expect_false(inherits(node_blip, "xml_missing"))
    expect_true(all(xml_attr(node_blip, "embed") %in% subset_rel$id))
  }
})


test_that("add multiple images in pptx", {
  x <- read_pptx()
  x <- add_slide(x)
  x <- ph_with(x = x, value = external_img(src = file1),
               location = ph_location(left = 0),
               use_loc_size = FALSE)
  x <- ph_with(x = x, value = external_img(src = file2),
               location = ph_location(left = 3),
               use_loc_size = FALSE)
  x <- ph_with(x = x, value = external_img(src = file3),
               location = ph_location(left = 6),
               use_loc_size = FALSE)
  pptx_file <- print(x, target = tempfile(fileext = ".pptx"))

  x <- read_pptx(path = pptx_file)
  slide <- x$slide$get_slide(x$cursor)

  rel_df <- slide$rel_df()

  subset_rel <- rel_df[grepl("^http://schemas(.*)image$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 1)
  if (nrow(subset_rel) > 0) {
    body <- slide$get()
    node_blip <- xml_find_all(body, "//a:blip")
    expect_length(node_blip, 3)
    expect_true(all(xml_attr(node_blip, "embed") %in% subset_rel$id))
  }
})

test_that("add svg in docx", {
  skip_if_not_installed("rsvg")
  x <- read_docx()
  x <- body_add_fpar(x, fpar(ext_svg))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  rel_df <- x$doc_obj$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)image$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 2)
  if (nrow(subset_rel) > 0) {
    body <- docx_body_xml(x)
    node_blip <- xml_find_first(body, "//a:blip")
    expect_false(inherits(node_blip, "xml_missing"))
    expect_true(all(xml_attr(node_blip, "embed") %in% subset_rel$id))
    node_svgblip <- xml_find_first(body, "//asvg:svgBlip")
    expect_false(inherits(node_svgblip, "xml_missing"))
    expect_true(all(xml_attr(node_svgblip, "embed") %in% subset_rel$id))
  }
})

test_that("add svg in pptx", {
  skip_if_not_installed("rsvg")
  x <- read_pptx()
  x <- add_slide(x)
  x <- ph_with(x, ext_svg, location = ph_location_type())
  filename <- print(x, target = tempfile(fileext = ".pptx"))

  x <- read_pptx(path = filename)
  slide <- x$slide$get_slide(x$cursor)

  rel_df <- slide$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)image$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 2)
  if (nrow(subset_rel) > 0) {
    node_blip <- xml_find_first(slide$get(), "//a:blip")
    expect_false(inherits(node_blip, "xml_missing"))
    expect_true(all(xml_attr(node_blip, "embed") %in% subset_rel$id))
    node_svgblip <- xml_find_first(slide$get(), "//asvg:svgBlip")
    expect_false(inherits(node_svgblip, "xml_missing"))
    expect_true(all(xml_attr(node_svgblip, "embed") %in% subset_rel$id))
  }
})

test_that("file size does not inflate with identical images", {
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  doc <- read_docx()
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
  file1 <- print(doc, target = tempfile(fileext = ".docx"))
  doc <- read_docx(path = file1)
  doc <- body_remove(doc)
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
  file2 <- print(doc, target = tempfile(fileext = ".docx"))
  expect_equal(file.size(file1), file.size(file2), tolerance = 10)
})

