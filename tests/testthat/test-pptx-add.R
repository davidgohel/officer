test_that("add wrong arguments", {
  doc <- read_pptx()
  expect_error(add_slide(doc, "Title and blah", "Office Theme"), fixed = TRUE)
  expect_error(add_slide(doc, "Title and Content", "Office Tddheme"), fixed = TRUE)
})


test_that("add simple elements into placeholder", {
  skip_if_not_installed("doconv")
  skip_if_not(doconv::msoffice_available())
  require(doconv)
  local_edition(3L)
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Two Content", master = "Office Theme")
  doc <- ph_with(doc, c(1L, 2L), location = ph_location_left())
  doc <- ph_with(doc, as.factor(c("rhhh", "vvvlllooo")), location = ph_location_right())
  doc <- add_slide(doc, layout = "Two Content", master = "Office Theme")
  doc <- ph_with(doc, c(pi, pi / 2), location = ph_location_left())
  doc <- ph_with(doc, c(TRUE, FALSE), location = ph_location_right())
  expect_snapshot_doc(x = doc, name = "pptx-add-simple", engine = "testthat")
})


test_that("add ggplot into placeholder", {
  skip_if_not_installed("doconv")
  skip_if_not_installed("ggplot2")
  skip_if_not(doconv::msoffice_available())
  require(doconv)
  require(ggplot2)
  local_edition(3L)
  doc <- read_pptx()
  doc <- add_slide(doc)
  gg_plot <- ggplot(data = iris) +
    geom_point(
      mapping = aes(Sepal.Length, Petal.Length),
      size = 3
    ) +
    theme_minimal()
  doc <- ph_with(
    x = doc, value = gg_plot,
    location = ph_location_type(type = "body"),
    bg = "transparent"
  )
  doc <- ph_with(
    x = doc, value = "graphic title",
    location = ph_location_type(type = "title")
  )
  expect_snapshot_doc(x = doc, name = "pptx-add-ggplot2", engine = "testthat")
})


test_that("add base plot into placeholder", {
  skip_if_not_installed("doconv")
  skip_if_not(doconv::msoffice_available())
  require(doconv)
  local_edition(3L)
  anyplot <- plot_instr(code = {
    barplot(1:5, col = 2:6)
  })
  doc <- read_pptx()
  doc <- add_slide(doc)
  doc <- ph_with(
    doc, anyplot,
    location = ph_location_fullsize(),
    bg = "#00000066", pointsize = 12
  )
  expect_snapshot_doc(x = doc, name = "pptx-add-barplot", engine = "testthat")
})


test_that("add unordered_list into placeholder", {
  skip_if_not_installed("doconv")
  skip_if_not(doconv::msoffice_available())
  require(doconv)
  local_edition(3L)
  ul1 <- unordered_list(
    level_list = c(0, 1, 1, 0, 0, 1, 1),
    str_list = c("List1", "Item 1", "Item 2", "", "List 2", "Option A", "Option B")
  )

  ul2 <- unordered_list(
    level_list = c(0, 1, 2, 0, 0, 1, 2) + 1,
    str_list = c("List1", "Item 1", "Item 2", "", "List 2", "Option A", "Option B"),
    style = fp_text_lite(color = "gray25")
  )

  doc <- read_pptx()
  doc <- add_slide(doc)
  doc <- ph_with(doc, ul1, location = ph_location_type())
  doc <- add_slide(doc)
  doc <- ph_with(doc, ul2, location = ph_location_type())
  expect_snapshot_doc(x = doc, name = "pptx-add-ul", engine = "testthat")
})


test_that("add block_list into placeholder", {
  skip_if_not_installed("doconv")
  skip_if_not(doconv::msoffice_available())
  require(doconv)
  local_edition(3L)
  fpt_blue_bold <- fp_text_lite(color = "#006699", bold = TRUE)
  fpt_red_italic <- fp_text_lite(color = "#C32900", italic = TRUE)
  value <- block_list(
    fpar(ftext("hello world", fpt_blue_bold)),
    fpar(
      ftext("hello", fpt_blue_bold), " ",
      ftext("world", fpt_red_italic)
    ),
    fpar(
      ftext("hello world", fpt_red_italic)
    )
  )

  doc <- read_pptx()
  doc <- add_slide(doc)
  doc <- ph_with(doc,
    value = value, location = ph_location_type(),
    level_list = c(1, 2, 3)
  )
  expect_snapshot_doc(x = doc, name = "pptx-add-blocklist", engine = "testthat")
})


test_that("add formatted par into placeholder", {
  bold_face <- shortcuts$fp_bold(font.size = 30)
  bold_redface <- update(bold_face, color = "red")

  fpar_ <- fpar(
    ftext("Hello ", prop = bold_face),
    ftext("World", prop = bold_redface),
    ftext(", how are you?", prop = bold_face)
  )

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, fpar_, location = ph_location_type(type = "body"))

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm[1, ]$text, "Hello World, how are you?")

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  cols <- xml_attr(xml_find_all(xmldoc, "//a:rPr/a:solidFill/a:srgbClr"), "val")
  expect_equal(cols, c("000000", "FF0000", "000000"))
  expect_equal(xml_attr(xml_find_all(xmldoc, "//a:rPr"), "b"), rep("1", 3))
})


test_that("add xml into placeholder", {
  xml_str <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr/>\n<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Hello world 1</a:t></a:r></a:p></p:txBody></p:sp>"
  library(xml2)
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- ph_with(doc, value = as_xml_document(xml_str), location = ph_location_type(type = "body"))
  doc <- ph_with(doc, value = as_xml_document(xml_str), location = ph_location(left = 1, top = 1, width = 3, height = 3))
  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 2)
  expect_equal(sm[1, ]$text, "Hello world 1")
  expect_equal(sm[2, ]$text, "Hello world 1")
})


test_that("slidelink shape", {
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, "Un titre 1", location = ph_location_type(type = "title"))
  doc <- ph_with(doc, "text 1", location = ph_location_type(type = "body"))
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, "Un titre 2", location = ph_location_type(type = "title"))
  doc <- on_slide(doc, index = 1)

  doc <- ph_slidelink(doc, type = "body", slide_index = 2)

  rel_df <- doc$slide$get_slide(1)$rel_df()

  slide_filename <- doc$slide$get_metadata()$name[2]

  expect_true(slide_filename %in% rel_df$target)
  row_num_ <- which(is.na(rel_df$target_mode) & rel_df$target %in% slide_filename)

  rid <- rel_df[row_num_, "id"]
  xpath_ <- sprintf("//p:sp[p:nvSpPr/p:cNvPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_)
  expect_false(inherits(node_, "xml_missing"))
})

test_that("hyperlink shape", {
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(x = doc, location = ph_location_type(type = "title"), value = "Un titre 1")
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(x = doc, location = ph_location_type(type = "title"), value = "Un titre 2")
  doc <- on_slide(doc, 1)
  doc <- ph_hyperlink(x = doc, type = "title", href = "https://cran.r-project.org")
  outputfile <- tempfile(fileext = ".pptx")
  print(doc, target = outputfile)

  doc <- read_pptx(outputfile)
  rel_df <- doc$slide$get_slide(1)$rel_df()

  expect_true("https://cran.r-project.org" %in% rel_df$target)
  row_num_ <- which(!is.na(rel_df$target_mode) & rel_df$target %in% "https://cran.r-project.org")

  rid <- rel_df[row_num_, "id"]
  xpath_ <- sprintf("//p:sp[p:nvSpPr/p:cNvPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_)
  expect_false(inherits(node_, "xml_missing"))
})

test_that("hyperlink image", {
  img.file <- file.path(R.home(component = "doc"), "html", "logo.jpg")

  x <- read_pptx()
  x <- add_slide(x)
  x <- ph_with(x, value = external_img(img.file), location = ph_location(newlabel = "logo"))
  x <- ph_hyperlink(x = x, ph_label = "logo", href = "https://cran.r-project.org")
  outputfile <- tempfile(fileext = ".pptx")
  print(x, target = outputfile)

  doc <- read_pptx(outputfile)
  rel_df <- doc$slide$get_slide(1)$rel_df()

  expect_true("https://cran.r-project.org" %in% rel_df$target)
})

test_that("img dims in pptx", {
  skip_on_os("windows")
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- ph_with(doc,
    value = external_img(img.file),
    location = ph_location(
      left = 1, top = 1,
      height = 1.06, width = 1.39
    )
  )
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)
  expect_equal(sm$cx, 1.39)
  expect_equal(sm$cy, 1.06)
  expect_equal(sm$offx, 1)
  expect_equal(sm$offy, 1)

  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- ph_with(doc,
    value = external_img(img.file),
    location = ph_location(
      left = 1, top = 1,
      height = 1.06, width = 1.39
    ),
    use_loc_size = TRUE
  )
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)
  expect_equal(sm$cx, 1.39, tolerance = .01)
  expect_equal(sm$cy, 1.06, tolerance = .01)
  expect_equal(sm$offx, 1)
  expect_equal(sm$offy, 1)
})

test_that("empty_content in pptx", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")
  doc <- ph_with(
    x = doc, value = empty_content(),
    location = ph_location(
      left = 0, top = 0,
      width = 2, height = 3, bg = "black"
    )
  )

  expect_equal(slide_summary(doc)$offy, 0)
  expect_equal(slide_summary(doc)$offx, 0)
  expect_equal(slide_summary(doc)$cy, 3)
  expect_equal(slide_summary(doc)$cx, 2)
})


test_that("pptx ph locations", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")

  doc <- ph_with(
    x = doc, value = "left",
    location = ph_location_left()
  )
  doc <- ph_with(
    x = doc, value = "right",
    location = ph_location_right()
  )
  doc <- ph_with(
    x = doc, value = "title",
    location = ph_location_type(type = "title")
  )
  doc <- ph_with(
    x = doc, value = "fullsize",
    location = ph_location_fullsize()
  )
  doc <- ph_with(
    x = doc, value = "from title",
    location = ph_location_template(
      left = 1, width = 2, height = 1, top = 4,
      type = "title", newlabel = "newlabel"
    )
  )

  layouts_info <- layout_properties(doc)


  title_xfrm <- layouts_info[layouts_info$name %in% "Two Content" &
    layouts_info$type %in% "title", c("offx", "offy", "cx", "cy")]
  side_xfrm <- layouts_info[layouts_info$name %in% "Two Content" &
    layouts_info$type %in% "body", c("offx", "offy", "cx", "cy")]
  full_xfrm <- as.data.frame(slide_size(doc))
  names(full_xfrm) <- c("cx", "cy")
  full_xfrm <- cbind(data.frame(offx = 0L, offy = 0L), full_xfrm)
  from_title_xfrm <- data.frame(offx = 1, offy = 4, cx = 2, cy = 1)
  theorical_xfrm <- rbind(
    side_xfrm,
    title_xfrm,
    full_xfrm,
    from_title_xfrm
  )

  all_xfrm <- xml_find_all(
    x = doc$slide$get_slide(1)$get(),
    xpath = "/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:xfrm"
  )
  offx <- xml_attr(xml_child(all_xfrm, "a:off"), "x")
  offx <- as.integer(offx) / 914400
  offy <- xml_attr(xml_child(all_xfrm, "a:off"), "y")
  offy <- as.integer(offy) / 914400
  cx <- xml_attr(xml_child(all_xfrm, "a:ext"), "cx")
  cx <- as.integer(cx) / 914400
  cy <- xml_attr(xml_child(all_xfrm, "a:ext"), "cy")
  cy <- as.integer(cy) / 914400

  observed_xfrm <- data.frame(offx = offx, offy = offy, cx = cx, cy = cy)
  expect_equivalent(observed_xfrm, theorical_xfrm)
})


test_that("pptx ph_location_type", {
  opts <- options(cli.num_colors = 1) # suppress colors for easier error message check
  on.exit(options(opts))

  x <- read_pptx()
  x <- add_slide(x, "Two Content")

  expect_no_error({
    ph_with(x, "correct ph type id", ph_location_type("body", type_idx = 1))
  })

  expect_warning({
    ph_with(x, "cannot supply id AND type_idx", ph_location_type("body", type_idx = 1, id = 1))
  }, regexp = "`id` is ignored if `type_idx` is provided", fixed = TRUE)

  expect_warning({
    ph_with(x, "id still working with warning to avoid breaking change", ph_location_type("body", id = 1))
  }, regexp = "The `id` argument in `ph_location_type()` is deprecated", fixed = TRUE)

  expect_error({
      ph_with(x, "out of range type id", ph_location_type("body", type_idx = 3)) # 3 does not exists => no error or warning
  }, regexp = "`type_idx` is out of range.", fixed = TRUE)

  expect_error({
    expect_warning({
    ph_with(x, "out of range type id", ph_location_type("body", id = 3)) # 3 does not exists => no error or warning
    }, regexp = " The `id` argument in `ph_location_type()` is deprecated", fixed = TRUE)
  }, regexp = "`id` is out of range.", fixed = TRUE)

  expect_error({
    ph_with(x, "type okay but not available in layout", ph_location_type("tbl")) # tbl not on layout
  }, regexp = "Found no placeholder of type", fixed = TRUE)

  expect_error({
    ph_with(x, "xxx is unknown type", ph_location_type("xxx"))
  }, regexp = 'type "xxx" is unknown', fixed = TRUE)

  expect_no_error({ # for complete coverage
    ph_with(x, " ph type position_right", ph_location_type("body", position_right = TRUE))
  })
})


test_that("pptx ph_location_id", {
  opts <- options(cli.num_colors = 1) # no colors for easier error message check
  on.exit(options(opts))

  # direct errors
  error_exp <- "`id` must be one number"
  expect_error(ph_location_id(id = 1:2), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = -1:1), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = c("A", "B")), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = c(NA, NA)), regex = error_exp, fixed = TRUE)

  error_exp <- "`id` must be a positive number"
  expect_error(ph_location_id(id = NULL), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = NA), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = NaN), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = character(0)), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = integer(0)), regex = error_exp, fixed = TRUE)

  expect_error(ph_location_id(id = "A"), regex = 'Cannot convert "A" to integer', fixed = TRUE)
  expect_error(ph_location_id(id = ""), regex = 'Cannot convert "" to integer', fixed = TRUE)
  expect_error(ph_location_id(id = Inf), regex = "Cannot convert Inf to integer", fixed = TRUE)
  expect_error(ph_location_id(id = -Inf), regex = "Cannot convert -Inf to integer", fixed = TRUE)

  error_exp <- "`id` must be a positive number"
  expect_error(ph_location_id(id = 0), regex = error_exp, fixed = TRUE)
  expect_error(ph_location_id(id = -1), regex = error_exp, fixed = TRUE)

  # downstream errors
  x <- read_pptx()
  x <- add_slide(x, "Comparison")

  expect_error(
    {
      ph_with(x, "id does not exist", ph_location_id(id = 1000))
    },
    "`id` 1000 does not exist",
    fixed = TRUE
  )

  # test for correct results
  expect_no_error({
    ids <- layout_properties(x, "Comparison")$id
    for (id in ids) {
      ph_with(x, paste("text:", id), ph_location_id(id, newlabel = paste("newlabel:", id)))
    }
  })
  nodes <- xml_find_all(
    x = x$slide$get_slide(1)$get(),
    xpath = "/p:sld/p:cSld/p:spTree/p:sp"
  )
  # text inside phs
  expect_true(all(xml_text(nodes) == paste("text:", ids)))
  # assigned shape names
  all_nvpr <- xml_find_all(nodes, "./p:nvSpPr/p:cNvPr")
  expect_true(all(xml_attr(all_nvpr, "name") == paste("newlabel:", ids)))
})


test_that("pptx ph labels", {
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content", "Office Theme")

  doc <- ph_with(
    x = doc, value = "elephant",
    location = ph_location_type(newlabel = "label1")
  )
  doc <- ph_with(
    x = doc, value = "elephant",
    location = ph_location_label(
      ph_label = "Date Placeholder 3",
      newlabel = "label2"
    )
  )
  doc <- ph_with(
    x = doc, value = "elephant",
    location = ph_location_left(newlabel = "label3")
  )
  doc <- ph_with(
    x = doc, value = "elephant",
    location = ph_location_right(newlabel = "label4")
  )

  all_nvpr <- xml_find_all(
    x = doc$slide$get_slide(1)$get(),
    xpath = "/p:sld/p:cSld/p:spTree/p:sp/p:nvSpPr/p:cNvPr"
  )
  expect_equal(
    xml_attr(all_nvpr, "name"),
    paste0("label", 1:4)
  )

  expect_error({
    doc <- ph_with(
      x = doc, value = "error if label does not exist",
      location = ph_location_label(ph_label = "xxx")
    )
  })
})



test_that("as_ph_location", {
  ref_names <- c("width", "height", "left", "top", "ph_label", "ph", "type", "rotation", "fld_id", "fld_type")
  l <- replicate(length(ref_names), "dummy", simplify = FALSE)
  df <- as.data.frame(l)
  names(df) <- ref_names

  expect_no_error({
    as_ph_location(df)
  })

  expect_error({
    as_ph_location(df[, -(1:2)])
  }, regexp = "missing column values:width,height", fixed = TRUE)

  expect_error({
    as_ph_location("wrong class supplied")
  }, regexp = "`x` must be a data frame", fixed = TRUE)
})


test_that("get_ph_loc", {
  x <- read_pptx()
  get_ph_loc(x, "Comparison", "Office Theme", type =  "body",
             position_right = TRUE, position_top = FALSE)

})


unlink("*.pptx")
