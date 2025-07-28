getncheck <- function(x, str) {
  child_ <- xml_child(x, str)
  expect_false(inherits(child_, "xml_missing"))
  child_
}

plot_with_unit_and_check <- function(
  value,
  body_add_fun,
  width = 6,
  height = 5,
  ...
) {
  .in_to_emu <- 914400
  .cm_to_emu <- 360000
  .mm_to_emu <- 36000

  x <- read_docx()
  x <- body_add_fun(x, value, width = width, height = height, ...)
  x <- body_add_fun(x, value, unit = "in", width = width, height = height, ...)
  x <- body_add_fun(x, value, unit = "cm", width = width, height = height, ...)
  x <- body_add_fun(x, value, unit = "mm", width = width, height = height, ...)
  x <- cursor_end(x)
  node <- docx_current_block_xml(x)

  expect_equal(
    as.numeric(
      xml_attr(xml_find_all(node, "//wp:extent"), "cx")
    ) /
      c(.in_to_emu, .in_to_emu, .cm_to_emu, .mm_to_emu),
    rep(width, 4)
  )
  expect_equal(
    as.numeric(
      xml_attr(xml_find_all(node, "//wp:extent"), "cy")
    ) /
      c(.in_to_emu, .in_to_emu, .cm_to_emu, .mm_to_emu),
    rep(height, 4)
  )
  # Non valid unit
  expect_error(
    body_add_fun(x, value, unit = "px", ...)
  )
  # Has "units="
  expect_error(
    body_add_fun(x, value, units = "cm", ...)
  )
  # Has multiple units
  expect_error(
    body_add_fun(x, value, unit = c("cm", "in", "mm"), ...)
  )
}

test_that("body_add_break", {
  x <- read_docx()
  x <- body_add_break(x)

  node <- docx_current_block_xml(x)
  expect_is(xml_child(node, "/w:r/w:br"), "xml_node")
})


test_that("body_end_sections", {
  x <- read_docx()
  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_end_section_landscape(x)

  node <- docx_current_block_xml(x)
  expect_false(inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing"))
  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false(inherits(ps, "xml_missing"))
  expect_equal(xml_attr(ps, "orient"), "landscape")

  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_add_par(x, "paragraph 2", style = "Normal")
  x <- body_end_section_columns(x)

  outfile <- tempfile(fileext = ".docx")
  print(x, target = outfile)
  x <- read_docx(outfile)

  node <- docx_current_block_xml(x)
  expect_false(inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing"))
  sect <- xml_child(node, "w:pPr/w:sectPr")

  expect_false(inherits(sect, "xml_missing"))
  expect_false(inherits(xml_child(sect, "w:cols"), "xml_missing"))

  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_add_par(x, "paragraph 2", style = "Normal")
  x <- body_end_section_columns_landscape(x)

  node <- docx_current_block_xml(x)
  expect_false(inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing"))

  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false(inherits(ps, "xml_missing"))
  expect_equal(xml_attr(ps, "orient"), "landscape")

  sect <- xml_child(node, "w:pPr/w:sectPr")
  expect_false(inherits(sect, "xml_missing"))
  expect_false(inherits(xml_child(sect, "w:cols"), "xml_missing"))

  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_add_par(x, "paragraph 2", style = "Normal")
  x <- body_end_section_portrait(x)

  outfile <- tempfile(fileext = ".docx")
  print(x, target = outfile)

  x <- read_docx(outfile)
  node <- docx_current_block_xml(x)
  expect_false(inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing"))

  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false(inherits(ps, "xml_missing"))
  expect_equal(xml_attr(ps, "orient"), "portrait")
})


test_that("body_add_toc", {
  x <- read_docx()
  x <- body_add_par(x, "paragraph 1")
  x <- body_add_toc(x)

  node <- docx_current_block_xml(x)

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal(xml_text(child_), "TOC \\o \"1-3\" \\h \\z \\u")

  x <- body_add_toc(x, style = "Normal")
  node <- docx_current_block_xml(x)

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal(xml_text(child_), "TOC \\h \\z \\t \"Normal\"")

  expect_output(print(block_toc(level = 2)), "TOC - max level: 2")
  expect_output(
    print(block_toc(level = 2, style = "Normal")),
    "TOC for style: Normal"
  )
  expect_output(
    print(block_toc(level = 2, seq_id = "tab")),
    "TOC for seq identifier: tab"
  )

  expect_match(
    to_wml(block_toc(seq_id = "tab")),
    "TOC \\\\h \\\\z \\\\c \"tab\""
  )
  expect_match(
    to_wml(block_toc(style = "Normal")),
    "TOC \\\\h \\\\z \\\\t \"Normal\""
  )
  expect_match(
    to_wml(block_toc(level = 2)),
    "TOC \\\\o \"1-2\" \\\\h \\\\z \\\\u"
  )
})

test_that("body_add_img", {
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  x <- read_docx()
  x <- body_add_img(x, img.file, width = 2.5, height = 1.3)

  node <- docx_current_block_xml(x)
  getncheck(node, "w:r/w:drawing")
})

test_that("body_add_img with units", {
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")

  plot_with_unit_and_check(img.file, body_add_img, height = 2.69, width = 3.53)
})

test_that("external_img add", {
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  x <- read_docx()
  x <- body_add_fpar(
    x = x,
    value = fpar(
      external_img(src = img.file, width = .3, height = .3)
    )
  )
  node <- docx_current_block_xml(x)
  getncheck(node, "w:r/w:drawing")
})


test_that("ggplot add", {
  testthat::skip_if_not(requireNamespace("ggplot2", quietly = TRUE))
  library("ggplot2")

  gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))
  x <- read_docx()
  x <- body_add_gg(x, value = gg_plot, style = "centered")
  x <- cursor_end(x)
  node <- docx_current_block_xml(x)
  getncheck(node, "w:r/w:drawing")
})

test_that("ggplot add with unit", {
  testthat::skip_if_not(requireNamespace("ggplot2", quietly = TRUE))
  library("ggplot2")

  gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))

  plot_with_unit_and_check(gg_plot, body_add_gg)
  plot_with_unit_and_check(gg_plot, body_add)
})

test_that("plot add with unit", {
  base_plot <- plot_instr(
    code = {
      barplot(1:5, col = 2:6)
    }
  )

  # Base plot errors for "mm" with default pointsize of 12.
  plot_with_unit_and_check(base_plot, body_add_plot, pointsize = 1)
  plot_with_unit_and_check(base_plot, body_add, pointsize = 1)
})

test_that("fpar add", {
  bold_face <- shortcuts$fp_bold(font.size = 20)
  bold_redface <- update(bold_face, color = "red")
  fpar_ <- fpar(
    ftext("This is a big ", prop = bold_face),
    ftext("text", prop = bold_redface)
  )
  fpar_ <- update(fpar_, fp_p = fp_par(text.align = "center"))
  x <- read_docx()
  x <- body_add_fpar(x, fpar_)

  node <- docx_current_block_xml(x)
  expect_equal(xml_text(node), "This is a big text")

  x <- read_docx()
  try(
    {
      x <- body_add_fpar(x, fpar_, style = "centered")
    },
    silent = TRUE
  )
  expect_is(x, "rdocx")
})
test_that("svg add", {
  skip_if_not_installed("rsvg")
  srcfile <- file.path(R.home("doc"), "html", "Rlogo.svg")
  x <- read_docx()
  x <- body_add_fpar(x, fpar(external_img(srcfile)))
  path <- print(x, target = tempfile(fileext = ".docx"))
  x <- read_docx(path = path)
  node <- docx_current_block_xml(x)
  reldf <- x$doc_obj$rel_df()
  relidsvg <- reldf[grepl("\\.svg$", reldf$target), "id"]
  relidpng <- reldf[grepl("\\.png$", reldf$target), "id"]
  node_blip <- xml_child(
    node,
    "w:r/w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip"
  )
  expect_equal(xml_attr(node_blip, "embed"), relidpng)
  node_svgblip <- xml_child(
    node,
    "w:r/w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip"
  )
  expect_equal(xml_attr(node_svgblip, "embed"), relidsvg)
})

test_that("add docx into docx", {

  file_1 <- tempfile(fileext = ".docx")
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")
  doc <- read_docx()
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39)
  print(doc, target = file_1)

  file_2 <- tempfile(fileext = ".docx")
  final_doc <- read_docx()
  doc <- body_add_docx(x = doc, src = file_1)
  print(doc, target = file_2)

  new_dir <- tempfile()
  unpack_folder(file_2, folder = new_dir)

  doc_parts <- read_xml(file.path(new_dir, "[Content_Types].xml"))
  doc_parts <- xml_find_all(
    doc_parts,
    "d1:Override[@ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']"
  )
  doc_parts <- xml_attr(doc_parts, "PartName")
  doc_parts <- basename(doc_parts)
  expect_equal(
    doc_parts[grepl("\\.docx$", doc_parts)],
    list.files(file.path(new_dir, "word"), pattern = "\\.docx$")
  )
})

test_that("Add comment at cursor position", {
  fp_bold <- fp_text_lite(bold = TRUE)
  fp_red <- fp_text_lite(color = "red")

  doc <- read_docx()
  doc <- body_add_par(doc, "This is a first Paragraph.")
  doc <- body_comment(
    doc,
    cmt = block_list("Comment on first par."),
    author = "Proofreader",
    date = Sys.Date()
  )
  doc <- body_add_fpar(
    doc,
    fpar("This is a second Paragraph. ", "This is a third Paragraph."),
    style = "Normal"
  )

  doc <- body_comment(
    doc,
    cmt = block_list(
      fpar(ftext("Comment on second par ...", fp_bold)),
      fpar(
        ftext("... with a second line.", fp_red)
      )
    ),
    author = "Proofreader 2",
    date = Sys.Date()
  )

  docx_file <- print(doc, target = tempfile(fileext = ".docx"))
  docx_dir <- tempfile()
  unpack_folder(docx_file, docx_dir)

  doc <- read_xml(file.path(docx_dir, "word/comments.xml"))
  comment1 <- xml_find_first(doc, "w:comment[@w:id='0']")
  comment2 <- xml_find_first(doc, "w:comment[@w:id='1']")

  expect_false(inherits(comment1, "xml_missing"))
  expect_false(inherits(comment2, "xml_missing"))

  expect_length(xml_children(comment1), 1)
  expect_length(xml_children(comment2), 2)
})


test_that("body_append_context - add content into docx", {
  doc <- read_docx()
  z <- body_append_start_context(doc)
  for (i in seq_len(10)) {
    write_elements_to_context(
      context = z,
      fpar("Hello World, ", i, fp_p = fp_par(word_style = "Normal")),
      fpar(run_pagebreak())
    )
  }
  doc <- body_append_stop_context(z)
  all_added_pars <- xml_find_all(
    docx_body_xml(doc),
    "w:body/w:p"
  )
  expect_length(all_added_pars, 20)
})


test_that("body_add R objects", {
  doc <- read_docx()
  doc <- body_add(doc, letters[1:5])
  doc <- body_add(doc, iris$Species[1:5])
  doc <- body_add(doc, 1:5)
  doc <- body_add(doc, fpar("hello"), style = "Normal")
  doc <- body_add(doc, head(iris), style = "table_template")
  docx_xml <- docx_body_xml(doc)

  doc_summary <- docx_summary(doc)

  txt <- doc_summary$text
  ref <- c("a", "b", "c", "d", "e", "setosa", "setosa", "setosa", "setosa",
           "setosa", "1", "2", "3", "4", "5", "hello", "Sepal.Length", "5.1",
           "4.9", "4.7", "4.6", "5.0", "5.4", "Sepal.Width", "3.5", "3.0",
           "3.2", "3.1", "3.6", "3.9", "Petal.Length", "1.4", "1.4", "1.3",
           "1.5", "1.4", "1.7", "Petal.Width", "0.2", "0.2", "0.2", "0.2",
           "0.2", "0.4", "Species", "setosa", "setosa", "setosa", "setosa",
           "setosa", "setosa")
  expect_equal(txt, ref)

  table_style_name <- doc_summary[doc_summary$content_type %in% "table cell", "style_name"]
  table_style_name <- unique(table_style_name)
  expect_equal(table_style_name, "table_template")
})
