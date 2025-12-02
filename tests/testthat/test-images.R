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

pic <- file.path(R.home("doc"), "html", "logo.jpg")
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
  x <- add_slide(x, "Title and Content")
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
  x <- add_slide(x, "Title and Content")
  x <- ph_with(
    x = x,
    value = external_img(src = file1),
    location = ph_location(left = 0),
    use_loc_size = FALSE
  )
  x <- ph_with(
    x = x,
    value = external_img(src = file2),
    location = ph_location(left = 3),
    use_loc_size = FALSE
  )
  x <- ph_with(
    x = x,
    value = external_img(src = file3),
    location = ph_location(left = 6),
    use_loc_size = FALSE
  )
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
  new_file <- print(x, target = tempfile(fileext = ".docx"))
  new_folder <- unpack_folder(new_file, tempfile())
  media_files <- list.files(file.path(new_folder, "word", "media"))
  expect_length(media_files, 2)
})

test_that("add svg in pptx", {
  skip_if_not_installed("rsvg")
  x <- read_pptx()
  x <- add_slide(x, "Title and Content")
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
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  doc <- read_docx()
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39)
  file1 <- print(doc, target = tempfile(fileext = ".docx"))
  doc <- read_docx(path = file1)
  doc <- body_remove(doc)
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39)
  file2 <- print(doc, target = tempfile(fileext = ".docx"))
  expect_equal(file.size(file1), file.size(file2), tolerance = 10)
})


# docx floating image tests ----

test_that("add floating image in docx with default params", {
  float_img <- floating_external_img(
    img.file,
    width = 2, height = 1.5,
    pos_x = 1, pos_y = 2
  )

  x <- read_docx()
  x <- body_add_fpar(x, fpar(float_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  rel_df <- x$doc_obj$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)image$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 1)

  if (nrow(subset_rel) > 0) {
    body <- docx_body_xml(x)

    # Check that anchor element exists (not inline)
    node_anchor <- xml_find_first(body, "//wp:anchor")
    expect_false(inherits(node_anchor, "xml_missing"))

    # Check inline does NOT exist
    node_inline <- xml_find_first(body, "//wp:inline")
    expect_true(inherits(node_inline, "xml_missing"))

    # Check image reference
    node_blip <- xml_find_first(body, "//a:blip")
    expect_false(inherits(node_blip, "xml_missing"))
    expect_true(all(xml_attr(node_blip, "embed") %in% subset_rel$id))

    # Check default wrap distances (0, 0, 0.125, 0.125 inches)
    # 0.125 inches = 114300 EMUs
    expect_equal(xml_attr(node_anchor, "distT"), "0")
    expect_equal(xml_attr(node_anchor, "distB"), "0")
    expect_equal(xml_attr(node_anchor, "distL"), "114300")
    expect_equal(xml_attr(node_anchor, "distR"), "114300")

    # Check default positioning (margin)
    node_pos_h <- xml_find_first(body, "//wp:positionH")
    expect_equal(xml_attr(node_pos_h, "relativeFrom"), "margin")

    node_pos_v <- xml_find_first(body, "//wp:positionV")
    expect_equal(xml_attr(node_pos_v, "relativeFrom"), "margin")

    # Check positions (1 inch = 914400 EMUs, 2 inches = 1828800 EMUs)
    pos_x_offset <- xml_text(xml_find_first(node_pos_h, "wp:posOffset"))
    pos_y_offset <- xml_text(xml_find_first(node_pos_v, "wp:posOffset"))
    expect_equal(pos_x_offset, "914400")   # 1 inch
    expect_equal(pos_y_offset, "1828800")  # 2 inches

    # Check default wrap (square, bothSides)
    node_wrap <- xml_find_first(body, "//wp:wrapSquare")
    expect_false(inherits(node_wrap, "xml_missing"))
    expect_equal(xml_attr(node_wrap, "wrapText"), "bothSides")

    # Check dimensions (2 inches = 1828800 EMUs, 1.5 inches = 1371600 EMUs)
    node_extent <- xml_find_first(body, "//wp:extent")
    expect_equal(xml_attr(node_extent, "cx"), "1828800")  # width
    expect_equal(xml_attr(node_extent, "cy"), "1371600")  # height
  }
})


test_that("add floating image with custom positioning", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    pos_x = 0.5, pos_y = 1.5,
    pos_h_from = "page",
    pos_v_from = "paragraph"
  )

  x <- read_docx()
  x <- body_add_fpar(x, fpar(float_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  body <- docx_body_xml(x)

  # Check positioning reference
  node_pos_h <- xml_find_first(body, "//wp:positionH")
  expect_equal(xml_attr(node_pos_h, "relativeFrom"), "page")

  node_pos_v <- xml_find_first(body, "//wp:positionV")
  expect_equal(xml_attr(node_pos_v, "relativeFrom"), "paragraph")

  # Check position values (0.5 inch = 457200 EMUs, 1.5 inches = 1371600 EMUs)
  pos_x_offset <- xml_text(xml_find_first(node_pos_h, "wp:posOffset"))
  pos_y_offset <- xml_text(xml_find_first(node_pos_v, "wp:posOffset"))
  expect_equal(pos_x_offset, "457200")
  expect_equal(pos_y_offset, "1371600")
})


test_that("add floating image with custom wrapping", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    wrap_type = "tight",
    wrap_side = "left"
  )

  x <- read_docx()
  x <- body_add_fpar(x, fpar(float_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  body <- docx_body_xml(x)

  # Check wrap type
  node_wrap <- xml_find_first(body, "//wp:wrapTight")
  expect_false(inherits(node_wrap, "xml_missing"))
  expect_equal(xml_attr(node_wrap, "wrapText"), "left")
})


test_that("add floating image with wrap topAndBottom", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    wrap_type = "topAndBottom"
  )

  x <- read_docx()
  x <- body_add_fpar(x, fpar(float_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  body <- docx_body_xml(x)

  # Check wrap type
  node_wrap <- xml_find_first(body, "//wp:wrapTopAndBottom")
  expect_false(inherits(node_wrap, "xml_missing"))
})


test_that("add floating image with custom distances", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    wrap_dist_top = 0.1,
    wrap_dist_bottom = 0.2,
    wrap_dist_left = 0.3,
    wrap_dist_right = 0.4
  )

  x <- read_docx()
  x <- body_add_fpar(x, fpar(float_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  body <- docx_body_xml(x)
  node_anchor <- xml_find_first(body, "//wp:anchor")

  # Check custom distances
  # 0.1 inch = 91440 EMUs
  # 0.2 inch = 182880 EMUs
  # 0.3 inch = 274320 EMUs
  # 0.4 inch = 365760 EMUs
  expect_equal(xml_attr(node_anchor, "distT"), "91440")
  expect_equal(xml_attr(node_anchor, "distB"), "182880")
  expect_equal(xml_attr(node_anchor, "distL"), "274320")
  expect_equal(xml_attr(node_anchor, "distR"), "365760")
})


test_that("add floating image with wrap none", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    wrap_type = "none"
  )

  x <- read_docx()
  x <- body_add_fpar(x, fpar(float_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  body <- docx_body_xml(x)

  # Check wrap type
  node_wrap <- xml_find_first(body, "//wp:wrapNone")
  expect_false(inherits(node_wrap, "xml_missing"))
})


test_that("add floating image with all custom params", {
  float_img <- floating_external_img(
    img.file,
    width = 2.5, height = 1.8,
    pos_x = 0.75, pos_y = 1.25,
    pos_h_from = "column",
    pos_v_from = "line",
    wrap_type = "through",
    wrap_side = "right",
    wrap_dist_top = 0.05,
    wrap_dist_bottom = 0.15,
    wrap_dist_left = 0.25,
    wrap_dist_right = 0.35
  )

  x <- read_docx()
  x <- body_add_fpar(x, fpar(float_img))
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  body <- docx_body_xml(x)
  node_anchor <- xml_find_first(body, "//wp:anchor")

  # Check positioning
  node_pos_h <- xml_find_first(body, "//wp:positionH")
  expect_equal(xml_attr(node_pos_h, "relativeFrom"), "column")
  pos_x_offset <- xml_text(xml_find_first(node_pos_h, "wp:posOffset"))
  expect_equal(pos_x_offset, "685800")  # 0.75 inch

  node_pos_v <- xml_find_first(body, "//wp:positionV")
  expect_equal(xml_attr(node_pos_v, "relativeFrom"), "line")
  pos_y_offset <- xml_text(xml_find_first(node_pos_v, "wp:posOffset"))
  expect_equal(pos_y_offset, "1143000")  # 1.25 inches

  # Check wrapping
  node_wrap <- xml_find_first(body, "//wp:wrapThrough")
  expect_false(inherits(node_wrap, "xml_missing"))
  expect_equal(xml_attr(node_wrap, "wrapText"), "right")

  # Check distances
  expect_equal(xml_attr(node_anchor, "distT"), "45720")   # 0.05 inch
  expect_equal(xml_attr(node_anchor, "distB"), "137160")  # 0.15 inch
  expect_equal(xml_attr(node_anchor, "distL"), "228600")  # 0.25 inch
  expect_equal(xml_attr(node_anchor, "distR"), "320040")  # 0.35 inch

  # Check dimensions (2.5 inches = 2286000 EMUs, 1.8 inches = 1645920 EMUs)
  node_extent <- xml_find_first(body, "//wp:extent")
  expect_equal(xml_attr(node_extent, "cx"), "2286000")
  expect_equal(xml_attr(node_extent, "cy"), "1645920")
})


# rtf floating image tests ----

test_that("add floating image in RTF with default params", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 0.75,
    pos_x = 0.5, pos_y = 1
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check shape structure exists
  expect_true(grepl("\\{\\\\shp", rtf_text))
  expect_true(grepl("\\{\\\\\\*\\\\shpinst", rtf_text))

  # Check shapeType = 75 (picture frame)
  expect_true(grepl("\\{\\\\sp\\{\\\\sn shapeType\\}\\{\\\\sv 75\\}\\}", rtf_text))

  # Check position (0.5 inch = 720 twips, 1 inch = 1440 twips)
  expect_true(grepl("\\\\shpleft720", rtf_text))
  expect_true(grepl("\\\\shptop1440", rtf_text))
  expect_true(grepl("\\\\shpright2160", rtf_text))  # left (720) + width (1440)
  expect_true(grepl("\\\\shpbottom2520", rtf_text)) # top (1440) + height (1080)

  # Check default positioning (margin)
  expect_true(grepl("\\\\shpbxmargin", rtf_text))
  expect_true(grepl("\\\\shpbymargin", rtf_text))

  # Check default wrap (square, both sides)
  expect_true(grepl("\\\\shpwr2", rtf_text))
  expect_true(grepl("\\\\shpwrk0", rtf_text))

  # Check image in front of text
  expect_true(grepl("\\\\shpfblwtxt0", rtf_text))

  # Check picture data exists
  expect_true(grepl("\\{\\\\sp\\{\\\\sn pib\\}", rtf_text))
  expect_true(grepl("\\{\\\\pict\\\\pngblip", rtf_text))

  # Check structure ends correctly
  expect_true(grepl("\\\\par\\}\\}\\}", rtf_text))
})


test_that("add floating image in RTF with custom positioning", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    pos_x = 1.5, pos_y = 2.5,
    pos_h_from = "page",
    pos_v_from = "paragraph"
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check position (1.5 inches = 2160 twips, 2.5 inches = 3600 twips)
  expect_true(grepl("\\\\shpleft2160", rtf_text))
  expect_true(grepl("\\\\shptop3600", rtf_text))

  # Check positioning reference
  expect_true(grepl("\\\\shpbxpage", rtf_text))
  expect_true(grepl("\\\\shpbypara", rtf_text))
})


test_that("add floating image in RTF with tight wrap left", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    pos_x = 1, pos_y = 1,
    wrap_type = "tight",
    wrap_side = "left"
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check wrap type (tight = 4)
  expect_true(grepl("\\\\shpwr4", rtf_text))

  # Check wrap side (left = 1)
  expect_true(grepl("\\\\shpwrk1", rtf_text))
})


test_that("add floating image in RTF with topAndBottom wrap", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    pos_x = 1, pos_y = 1,
    wrap_type = "topAndBottom"
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check wrap type (topAndBottom = 1)
  expect_true(grepl("\\\\shpwr1", rtf_text))
})


test_that("add floating image in RTF with through wrap right", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    pos_x = 1, pos_y = 1,
    wrap_type = "through",
    wrap_side = "right"
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check wrap type (through = 5)
  expect_true(grepl("\\\\shpwr5", rtf_text))

  # Check wrap side (right = 2)
  expect_true(grepl("\\\\shpwrk2", rtf_text))
})


test_that("add floating image in RTF with none wrap", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    pos_x = 1, pos_y = 1,
    wrap_type = "none"
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check wrap type (none = 3)
  expect_true(grepl("\\\\shpwr3", rtf_text))
})


test_that("add floating image in RTF with custom wrap distances", {
  float_img <- floating_external_img(
    img.file,
    width = 1, height = 1,
    pos_x = 1, pos_y = 1,
    wrap_dist_top = 0.1,
    wrap_dist_bottom = 0.2,
    wrap_dist_left = 0.15,
    wrap_dist_right = 0.25
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check wrap distances in EMUs
  # 0.1 inch = 91440 EMUs
  # 0.2 inch = 182880 EMUs
  # 0.15 inch = 137160 EMUs
  # 0.25 inch = 228600 EMUs
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dxWrapDistLeft\\}\\{\\\\sv 137160\\}\\}", rtf_text))
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dxWrapDistRight\\}\\{\\\\sv 228600\\}\\}", rtf_text))
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dyWrapDistTop\\}\\{\\\\sv 91440\\}\\}", rtf_text))
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dyWrapDistBottom\\}\\{\\\\sv 182880\\}\\}", rtf_text))
})


test_that("add floating image in RTF with all custom params", {
  float_img <- floating_external_img(
    img.file,
    width = 2, height = 1.5,
    pos_x = 0.75, pos_y = 1.25,
    pos_h_from = "column",
    pos_v_from = "page",
    wrap_type = "square",
    wrap_side = "largest",
    wrap_dist_top = 0.05,
    wrap_dist_bottom = 0.1,
    wrap_dist_left = 0.15,
    wrap_dist_right = 0.2
  )

  doc <- rtf_doc()
  doc <- rtf_add(doc, fpar(float_img))
  rtf_file <- print(doc, target = tempfile(fileext = ".rtf"))

  rtf_content <- readLines(rtf_file, warn = FALSE)
  rtf_text <- paste(rtf_content, collapse = "")

  # Check position (0.75 inch = 1080 twips, 1.25 inch = 1800 twips)
  expect_true(grepl("\\\\shpleft1080", rtf_text))
  expect_true(grepl("\\\\shptop1800", rtf_text))
  # Right = left + width = 1080 + 2880 = 3960
  # Bottom = top + height = 1800 + 2160 = 3960
  expect_true(grepl("\\\\shpright3960", rtf_text))
  expect_true(grepl("\\\\shpbottom3960", rtf_text))

  # Check positioning reference
  expect_true(grepl("\\\\shpbxcolumn", rtf_text))
  expect_true(grepl("\\\\shpbypage", rtf_text))

  # Check wrap type (square = 2) and side (largest = 3)
  expect_true(grepl("\\\\shpwr2", rtf_text))
  expect_true(grepl("\\\\shpwrk3", rtf_text))

  # Check wrap distances
  # 0.05 inch = 45720 EMUs
  # 0.1 inch = 91440 EMUs
  # 0.15 inch = 137160 EMUs
  # 0.2 inch = 182880 EMUs
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dyWrapDistTop\\}\\{\\\\sv 45720\\}\\}", rtf_text))
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dyWrapDistBottom\\}\\{\\\\sv 91440\\}\\}", rtf_text))
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dxWrapDistLeft\\}\\{\\\\sv 137160\\}\\}", rtf_text))
  expect_true(grepl("\\{\\\\sp\\{\\\\sn dxWrapDistRight\\}\\{\\\\sv 182880\\}\\}", rtf_text))

  # Check picture dimensions in twips (2 inch = 2880 twips, 1.5 inch = 2160 twips)
  expect_true(grepl("\\\\picwgoal2880", rtf_text))
  expect_true(grepl("\\\\pichgoal2160", rtf_text))
})


# plot_in_png tests ----

test_that("plot_in_png with ggplot object", {
  skip_if_not_installed("ggplot2")

  gg <- ggplot2::ggplot(mtcars, ggplot2::aes(x = mpg, y = hp)) +
    ggplot2::geom_point()

  png_file <- plot_in_png(
    ggobj = gg,
    width = 5,
    height = 4,
    res = 72,
    units = "in"
  )

  expect_true(file.exists(png_file))
  expect_match(png_file, "\\.png$")
  expect_true(file.size(png_file) > 0)
})

test_that("plot_in_png with code expression", {
  png_file <- plot_in_png(
    code = {
      plot(1:10, 1:10)
    },
    width = 5,
    height = 4,
    res = 72,
    units = "in"
  )

  expect_true(file.exists(png_file))
  expect_match(png_file, "\\.png$")
  expect_true(file.size(png_file) > 0)
})

test_that("plot_in_png with custom path", {
  custom_path <- tempfile(fileext = ".png")

  png_file <- plot_in_png(
    code = {
      barplot(1:5)
    },
    width = 4,
    height = 3,
    res = 96,
    units = "in",
    path = custom_path
  )

  expect_equal(png_file, custom_path)
  expect_true(file.exists(custom_path))
})


# as_base64 and from_base64 tests ----

test_that("as_base64 with multiple values", {
  input <- c("hello", "world", "test")
  result <- as_base64(input)

  expect_type(result, "character")
  expect_length(result, 3)
  expect_false(any(is.na(result)))
})

test_that("as_base64 with NA values", {
  input <- c("hello", NA_character_, "world")
  result <- as_base64(input)

  expect_length(result, 3)
  expect_equal(result[1], as_base64("hello"))
  expect_true(is.na(result[2]))
  expect_equal(result[3], as_base64("world"))
})

test_that("as_base64 with invalid input", {
  expect_error(as_base64(123), "'x' must be a character vector")
  expect_error(as_base64(list("a", "b")), "'x' must be a character vector")
})

test_that("from_base64 with multiple values", {
  original <- c("hello", "world", "test")
  encoded <- as_base64(original)
  decoded <- from_base64(encoded)

  expect_equal(decoded, original)
})

test_that("from_base64 with NA values", {
  encoded <- c(as_base64("hello"), NA_character_, as_base64("world"))
  result <- from_base64(encoded)

  expect_length(result, 3)
  expect_equal(result[1], "hello")
  expect_true(is.na(result[2]))
  expect_equal(result[3], "world")
})

test_that("from_base64 with invalid input type", {
  expect_error(from_base64(123), "'x' must be a character vector")
})

test_that("from_base64 with invalid base64 string", {
  expect_error(from_base64("not_valid_base64!!!"), "Failed to decode Base64 element")
})


# base64_to_image tests ----

test_that("base64_to_image converts data URI to image file", {
  img1 <- file.path(R.home("doc"), "html", "logo.jpg")
  img2 <- file.path(R.home("doc"), "html", "Rlogo.svg")
  base64_str <- image_to_base64(c(img1, img2))

  output_files <- c(
    tempfile(fileext = ".jpg"),
    tempfile(fileext = ".svg")
  )
  result <- base64_to_image(base64_str, output_files = output_files)

  expect_equal(result, output_files)
  expect_true(all(file.exists(output_files)))
  expect_true(all(file.size(output_files) > 0))
})


# image_to_base64 error handling tests ----

test_that("image_to_base64 with multiple files", {
  img1 <- file.path(R.home("doc"), "html", "logo.jpg")
  img2 <- file.path(R.home("doc"), "html", "Rlogo.svg")

  result <- image_to_base64(c(img1, img2))

  expect_type(result, "character")
  expect_length(result, 2)
  expect_match(result[1], "^data:image/jpeg;base64,")
  expect_match(result[2], "^data:image/svg\\+xml;base64,")
})

test_that("image_to_base64 with unknown format", {
  temp_file <- tempfile(fileext = ".xyz")
  writeLines("test", temp_file)

  expect_error(
    image_to_base64(temp_file),
    "Unknown image\\(s\\) format"
  )
})

test_that("image_to_base64 with non-existent file", {
  fake_file <- tempfile(fileext = ".png")

  expect_error(
    image_to_base64(fake_file),
    "File\\(s\\) not found"
  )
})

test_that("image_to_base64 with multiple non-existent files", {
  fake1 <- tempfile(fileext = ".png")
  fake2 <- tempfile(fileext = ".jpg")

  expect_error(
    image_to_base64(c(fake1, fake2)),
    "File\\(s\\) not found"
  )
})
