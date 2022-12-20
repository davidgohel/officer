getncheck <- function(x, str){
  child_ <- xml_child(x, str)
  expect_false( inherits(child_, "xml_missing") )
  child_
}

test_that("body_add_break", {
  x <- read_docx()
  x <- body_add_break(x)

  node <- docx_current_block_xml(x)
  expect_is( xml_child(node, "/w:r/w:br"), "xml_node" )
})


test_that("body_end_sections", {

  x <- read_docx()
  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_end_section_landscape(x)

  node <- docx_current_block_xml(x)
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )
  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false( inherits(ps, "xml_missing") )
  expect_equal( xml_attr(ps, "orient"), "landscape")

  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_add_par(x, "paragraph 2", style = "Normal")
  x <- body_end_section_columns(x)

  outfile <- tempfile(fileext = ".docx")
  print(x, target = outfile)
  x <- read_docx(outfile)

  node <- docx_current_block_xml(x)
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )
  sect <- xml_child(node, "w:pPr/w:sectPr")

  expect_false( inherits(sect, "xml_missing") )
  expect_false( inherits(xml_child(sect, "w:cols"), "xml_missing") )

  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_add_par(x, "paragraph 2", style = "Normal")
  x <- body_end_section_columns_landscape(x)

  node <- docx_current_block_xml(x)
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )

  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false( inherits(ps, "xml_missing") )
  expect_equal( xml_attr(ps, "orient"), "landscape")

  sect <- xml_child(node, "w:pPr/w:sectPr")
  expect_false( inherits(sect, "xml_missing") )
  expect_false( inherits(xml_child(sect, "w:cols"), "xml_missing") )

  x <- body_add_par(x, "paragraph 1", style = "Normal")
  x <- body_add_par(x, "paragraph 2", style = "Normal")
  x <- body_end_section_portrait(x)

  outfile <- tempfile(fileext = ".docx")
  print(x, target = outfile)

  x <- read_docx(outfile)
  node <- docx_current_block_xml(x)
  expect_false( inherits(xml_child(node, "w:pPr/w:sectPr"), "xml_missing") )

  ps <- xml_child(node, "w:pPr/w:sectPr/w:pgSz")
  expect_false( inherits(ps, "xml_missing") )
  expect_equal( xml_attr(ps, "orient"), "portrait")
})


test_that("body_add_toc", {

  x <- read_docx()
  x <- body_add_par(x, "paragraph 1")
  x <- body_add_toc(x)

  node <- docx_current_block_xml(x)

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal( xml_text(child_), "TOC \\o \"1-3\" \\h \\z \\u" )


  x <- body_add_toc(x, style = "Normal")
  node <- docx_current_block_xml(x)

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")

  child_ <- getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal( xml_text(child_), "TOC \\h \\z \\t \"Normal;1\"" )

})

test_that("body_add_img", {

  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  x <- read_docx()
  x <- body_add_img(x, img.file, width=2.5, height=1.3)

  node <- docx_current_block_xml(x)
  getncheck(node, "w:r/w:drawing")
})

test_that("external_img add", {
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
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

  gg_plot <- ggplot(data = iris ) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))
  x <- read_docx()
  x <- body_add_gg(x, value = gg_plot, style = "centered" )
  x <- cursor_end(x)
  node <- docx_current_block_xml(x)
  getncheck(node, "w:r/w:drawing")
})

test_that("fpar add", {
  bold_face <- shortcuts$fp_bold(font.size = 20)
  bold_redface <- update(bold_face, color = "red")
  fpar_ <- fpar(ftext("This is a big ", prop = bold_face),
                ftext("text", prop = bold_redface ) )
  fpar_ <- update(fpar_, fp_p = fp_par(text.align = "center"))
  x <- read_docx()
  x <- body_add_fpar(x, fpar_)

  node <- docx_current_block_xml(x)
  expect_equal(xml_text(node), "This is a big text" )

  x <- read_docx()
  try({x <- body_add_fpar(x, fpar_, style = "centered")}, silent = TRUE)
  expect_is(x, "rdocx")

})
test_that("svg add", {
  skip_if_not_installed("rsvg")
  srcfile <- file.path( R.home("doc"), "html", "Rlogo.svg" )
  x <- read_docx()
  x <- body_add_fpar(x, fpar(external_img(srcfile)))
  path <- print(x, target = tempfile(fileext = ".docx"))
  x <- read_docx(path = path)
  node <- docx_current_block_xml(x)
  reldf <- x$doc_obj$rel_df()
  relidsvg <- reldf[grepl("\\.svg$", reldf$target), "id"]
  relidpng <- reldf[grepl("\\.png$", reldf$target), "id"]
  node_blip <- xml_child(node, "w:r/w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip")
  expect_equal(xml_attr(node_blip, "embed"), relidpng)
  node_svgblip <- xml_child(node, "w:r/w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip")
  expect_equal(xml_attr(node_svgblip, "embed"), relidsvg)
})

test_that("add docx into docx", {

  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  doc <- read_docx()
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39 )
  print(doc, target = "external_file.docx")

  final_doc <- read_docx()
  doc <- body_add_docx(x = doc, src = "external_file.docx" )
  print(doc, target = "final.docx")

  new_dir <- tempfile()
  unpack_folder("final.docx", folder = new_dir)

  doc_parts <- read_xml(file.path(new_dir, "[Content_Types].xml"))
  doc_parts <- xml_find_all(doc_parts, "d1:Override[@ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']")
  doc_parts <- xml_attr(doc_parts, "PartName")
  doc_parts <- basename(doc_parts)
  expect_equal(doc_parts[grepl("\\.docx$", doc_parts)],
               list.files(file.path(new_dir, "word"), pattern = "\\.docx$") )
})


unlink("*.docx")
unlink("*.emf")

img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
fpt_blue_bold <- fp_text(color = "#006699", bold = TRUE)
fpt_red_italic <- fp_text(color = "#C32900", italic = TRUE)
bl <- block_list(
  fpar(ftext("hello world", fpt_blue_bold)),
  fpar(ftext("hello", fpt_blue_bold), " ",
       ftext("world", fpt_red_italic)),
  fpar(
    ftext("hello world", fpt_red_italic),
    external_img(
      src = img.file, height = 1.06, width = 1.39)))

anyplot <- plot_instr(code = {
  col <- c("#440154FF", "#443A83FF", "#31688EFF",
                      "#21908CFF", "#35B779FF", "#8FD744FF", "#FDE725FF")
                      barplot(1:7, col = col, yaxt="n")
})

test_that("visual testing", {
  local_edition(3)
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)

  x <- read_docx()
  # add text and a table ----
  x <- body_add_par(x, "Hello World")
  x <- body_add_par(x, "Hello title", style = "heading 1")
  x <- body_add_par(x, "Hello title", style = "heading 2")
  x <- body_add_table(x, head(cars))
  x <- body_add_par(x, "Hello base plot", style = "heading 2")
  x <- body_add_plot(x, anyplot)
  x <- body_add_par(x, "Hello fpars", style = "heading 2")
  x <- body_add_blocks(x = x, blocks = bl)

  expect_snapshot_doc(x = x, name = "docx-elements", engine = "testthat")
})

test_that("body_add visual testing", {
  local_edition(3)
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)

  x <- read_docx()
  # add text and a table ----
  x <- body_add(x, "Hello World")
  x <- body_add(x, "Hello title", style = "heading 1")
  x <- body_add(x, "Hello title", style = "heading 2")
  x <- body_add(x, head(cars))
  x <- body_add(x, "Hello base plot", style = "heading 2")
  x <- body_add(x, anyplot)
  x <- body_add(x, "Hello fpars", style = "heading 2")
  x <- body_add(x = x, bl)

  expect_snapshot_doc(x = x, name = "body_add-elements", engine = "testthat")
})
