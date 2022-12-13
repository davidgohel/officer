img.file <- file.path(R.home(component = "doc"), "html", "logo.jpg")
svg_file <- file.path(R.home(component = "doc"), "html", "Rlogo.svg")

ext_img <- external_img(img.file)
ext_svg <- external_img(svg_file)


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
