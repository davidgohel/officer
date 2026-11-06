test_that("add wrong arguments", {
  x <- read_pptx()
  expect_error(add_slide(x, "Title and XXXX", "Office Theme"), fixed = TRUE)
  expect_error(add_slide(x, "Title and Content", "Office XXXX"), fixed = TRUE)
})


test_that("default layout", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()
  expect_warning(
    add_slide(x),
    regexp = "Calling `add_slide()` without specifying a `layout` is deprecated",
    fixed = TRUE
  )
  expect_no_error(expect_warning(add_slide(x)))

  x <- layout_default(x, "Title Slide")
  expect_no_error(add_slide(x))
  expect_no_error(add_slide(x, layout = "Two Content"))
})


test_that("master is inferred", {
  opts <- options(cli.num_colors = 1) # suppress colors for error message check
  on.exit(options(opts))

  x <- read_pptx()
  x <- add_slide(x, "Title Slide")
  x <- add_slide(x, "Title and Content")
  expect_equal(length(x), 2)

  file <- test_path("docs_dir", "test-layout-unique-and-dupe.pptx")
  x <- read_pptx(file)
  x <- add_slide(x, "Title Slide", "Master_1")
  x <- add_slide(x, "Title Slide") # name is unique so master=NULL works
  expect_error(
    add_slide(x, "Title and Content"),
    regex = "Layout \"Title and Content\" exists in more than one master"
  )
  x <- add_slide(x, "Title and Content", "Master_1")
  expect_equal(length(x), 3)
})


# text, level, color, bold and italic of the paragraphs of a slide
slide_paragraphs <- function(doc, slide_idx = 1) {
  xmldoc <- doc$slide$get_slide(id = slide_idx)$get()
  pars <- xml_find_all(xmldoc, "//p:txBody/a:p")
  first_run <- xml_find_first(pars, ".//a:rPr")
  data.frame(
    text = xml_text(pars),
    level = xml_attr(xml_find_first(pars, "a:pPr"), "lvl"),
    color = xml_attr(xml_find_first(first_run, "a:solidFill/a:srgbClr"), "val"),
    bold = xml_attr(first_run, "b"),
    italic = xml_attr(first_run, "i")
  )
}

test_that("simple elements into placeholder", {
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Two Content", master = "Office Theme")
  doc <- ph_with(doc, c(1L, 2L), location = ph_location_left())
  doc <- ph_with(
    doc,
    as.factor(c("rhhh", "vvvlllooo")),
    location = ph_location_right()
  )
  doc <- add_slide(doc, layout = "Two Content", master = "Office Theme")
  doc <- ph_with(doc, c(pi, pi / 2), location = ph_location_left())
  doc <- ph_with(doc, c(TRUE, FALSE), location = ph_location_right())

  sm <- slide_summary(doc, index = 1)
  sm <- sm[order(sm$offx), ]
  expect_equal(sm$text, c("12", "rhhhvvvlllooo"))
  sm <- slide_summary(doc, index = 2)
  sm <- sm[order(sm$offx), ]
  # format() pads TRUE to the width of FALSE
  expect_equal(trimws(sm$text), c("3.1415931.570796", "TRUEFALSE"))
})


test_that("date formats in placeholder", {
  old_opts <- options(officer.date_format = "%D")
  on.exit(options(old_opts))
  my_date <- as.Date("2025-09-02")

  x <- read_pptx()
  x <- add_slide(x, layout = "Two Content", master = "Office Theme")
  x <- ph_with(x, my_date, ph_location_type("body", type_idx = 1))
  options(officer.date_format = NA) # NA should be okay as well
  x <- ph_with(x, my_date, ph_location_type("body", type_idx = 2))
  x <- ph_with(x, my_date, ph_location_type("dt"))
  x <- ph_with(x, my_date, ph_location_type("ftr"), date_format = "%d.%m.%Y")

  sm <- slide_summary(x)
  expect_equal(sm$text[sm$type == "body"], c("09/02/25", "2025-09-02"))
  expect_equal(sm$text[sm$type == "dt"], "2025-09-02")
  expect_equal(sm$text[sm$type == "ftr"], "02.09.2025")
})

test_that("add factor into placeholder", {
  my_factor <- factor(
    c("Level A", "Level B", "Level C"),
    levels = c("Level A", "Level B", "Level C")
  )

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, my_factor, location = ph_location_type(type = "body"))

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "Level ALevel BLevel C")
})

test_that("add date into placeholder", {
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(
    doc,
    as.Date("2024-12-31", "%Y-%m-%d"),
    location = ph_location_type(type = "body")
  )

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "2024-12-31")
})

test_that("add unordered_list into placeholder", {
  my_ul <- unordered_list(
    level_list = c(1, 2, 1),
    str_list = c("Item 1", "Subitem", "Item 2")
  )

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, my_ul, location = ph_location_type(type = "body"))

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm$text, "Item 1SubitemItem 2")
})

test_that("add plot_instr with custom dimensions", {
  anyplot <- plot_instr(code = {
    barplot(1:5, col = 2:6)
  })

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(
    doc,
    anyplot,
    location = ph_location(left = 1, top = 1, width = 4, height = 3)
  )

  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm$cx, 4)
  expect_equal(sm$cy, 3)
})


test_that("ggplot into placeholder", {
  skip_if_not_installed("ggplot2")
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content")
  gg_plot <- ggplot2::ggplot(data = iris) +
    ggplot2::geom_point(
      mapping = ggplot2::aes(Sepal.Length, Petal.Length),
      size = 3
    )
  doc <- ph_with(
    x = doc,
    value = gg_plot,
    location = ph_location_type(type = "body"),
    bg = "transparent"
  )
  doc <- ph_with(
    x = doc,
    value = "graphic title",
    location = ph_location_type(type = "title")
  )

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  expect_length(xml_find_all(xmldoc, "//p:pic"), 1L)
  sm <- slide_summary(doc)
  body <- layout_properties(doc, "Title and Content", "Office Theme")
  body <- body[body$type == "body", ]
  expect_equal(sm$cx[sm$type == "body"], body$cx)
  expect_equal(sm$cy[sm$type == "body"], body$cy)
  expect_equal(sm$text[sm$type == "title"], "graphic title")
})


test_that("unordered_list levels and style in placeholder", {
  strings <- c(
    "List1",
    "Item 1",
    "Item 2",
    "",
    "List 2",
    "Option A",
    "Option B"
  )
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(
    doc,
    unordered_list(level_list = c(0, 1, 1, 0, 0, 1, 1), str_list = strings),
    location = ph_location_type()
  )
  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(
    doc,
    unordered_list(
      level_list = c(1, 2, 3, 1, 1, 2, 3),
      str_list = strings,
      style = fp_text_lite(color = "gray25")
    ),
    location = ph_location_type()
  )

  no_style <- slide_paragraphs(doc, slide_idx = 1)
  expect_equal(no_style$text, strings)
  expect_true(all(is.na(no_style$level)))
  expect_true(all(is.na(no_style$color)))

  styled <- slide_paragraphs(doc, slide_idx = 2)
  expect_equal(styled$text, strings)
  expect_equal(styled$level, c(NA, "1", "2", NA, NA, "1", "2"))
  expect_equal(styled$color, rep("404040", length(strings)))
})


test_that("add block_list into placeholder", {
  fpt_blue_bold <- fp_text_lite(color = "#006699", bold = TRUE)
  fpt_red_italic <- fp_text_lite(color = "#C32900", italic = TRUE)
  value <- block_list(
    fpar(ftext("hello world", fpt_blue_bold)),
    fpar(
      ftext("hello", fpt_blue_bold),
      " ",
      ftext("world", fpt_red_italic)
    ),
    fpar(
      ftext("hello world", fpt_red_italic)
    )
  )

  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(
    doc,
    value = value,
    location = ph_location_type(),
    level_list = c(1, 2, 3)
  )

  pars <- slide_paragraphs(doc)
  expect_equal(pars$text, c("hello world", "hello world", "hello world"))
  expect_equal(pars$level, c(NA, "1", "2"))

  # all the runs of the second paragraph keep their own format
  xmldoc <- doc$slide$get_slide(id = 1)$get()
  runs <- xml_find_all(xmldoc, "//p:txBody/a:p[2]/a:r/a:rPr")
  expect_equal(
    xml_attr(xml_find_first(runs, "a:solidFill/a:srgbClr"), "val"),
    c("006699", "C32900")
  )
  expect_equal(xml_attr(runs, "b"), c("1", NA))
  expect_equal(xml_attr(runs, "i"), c(NA, "1"))

  expect_equal(pars$color[c(1, 3)], c("006699", "C32900"))
  expect_equal(pars$bold[c(1, 3)], c("1", NA))
  expect_equal(pars$italic[c(1, 3)], c(NA, "1"))
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
  doc <- ph_with(
    doc,
    value = as_xml_document(xml_str),
    location = ph_location_type(type = "body")
  )
  doc <- ph_with(
    doc,
    value = as_xml_document(xml_str),
    location = ph_location(left = 1, top = 1, width = 3, height = 3)
  )
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
  row_num_ <- which(
    is.na(rel_df$target_mode) & rel_df$target %in% slide_filename
  )

  rid <- rel_df[row_num_, "id"]
  xpath_ <- sprintf("//p:sp[p:nvSpPr/p:cNvPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_)
  expect_false(inherits(node_, "xml_missing"))
})

test_that("hyperlink shape", {
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(
    x = doc,
    location = ph_location_type(type = "title"),
    value = "Un titre 1"
  )
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(
    x = doc,
    location = ph_location_type(type = "title"),
    value = "Un titre 2"
  )
  doc <- on_slide(doc, 1)
  doc <- ph_hyperlink(
    x = doc,
    type = "title",
    href = "https://cran.r-project.org"
  )
  outputfile <- tempfile(fileext = ".pptx")
  print(doc, target = outputfile)

  doc <- read_pptx(outputfile)
  rel_df <- doc$slide$get_slide(1)$rel_df()

  expect_true("https://cran.r-project.org" %in% rel_df$target)
  row_num_ <- which(
    !is.na(rel_df$target_mode) & rel_df$target %in% "https://cran.r-project.org"
  )

  rid <- rel_df[row_num_, "id"]
  xpath_ <- sprintf("//p:sp[p:nvSpPr/p:cNvPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_)
  expect_false(inherits(node_, "xml_missing"))
})

test_that("hyperlink image", {
  img.file <- file.path(R.home(component = "doc"), "html", "logo.jpg")

  x <- read_pptx()
  x <- add_slide(x, "Title and Content")
  x <- ph_with(
    x,
    value = external_img(img.file),
    location = ph_location(newlabel = "logo")
  )
  x <- ph_hyperlink(
    x = x,
    ph_label = "logo",
    href = "https://cran.r-project.org"
  )
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
  doc <- ph_with(
    doc,
    value = external_img(img.file),
    location = ph_location(
      left = 1,
      top = 1,
      height = 1.06,
      width = 1.39
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
  doc <- ph_with(
    doc,
    value = external_img(img.file),
    location = ph_location(
      left = 1,
      top = 1,
      height = 1.06,
      width = 1.39
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
    x = doc,
    value = empty_content(),
    location = ph_location(
      left = 0,
      top = 0,
      width = 2,
      height = 3,
      bg = "black"
    )
  )

  expect_equal(slide_summary(doc)$offy, 0)
  expect_equal(slide_summary(doc)$offx, 0)
  expect_equal(slide_summary(doc)$cy, 3)
  expect_equal(slide_summary(doc)$cx, 2)
})


unlink("*.pptx")
