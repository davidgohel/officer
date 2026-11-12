skip_if_not_installed("equatags", minimum_version = "0.4.0")

test_that("run_equation() checks its arguments", {
  expect_s3_class(run_equation("x^2"), c("run_equation", "run"))
  expect_error(run_equation(c("x", "y")), "single LaTeX string")
  expect_error(run_equation(NA_character_), "single LaTeX string")
  expect_error(run_equation("x", prop = fp_par()), "fp_text")
  expect_identical(format(run_equation("x^2")), "x^2")
})

test_that("an equation in a sentence is an inline equation", {
  xml <- to_wml(fpar("Euler: ", run_equation("\\sum_{i=1}^{n} x_i"), "."))
  expect_match(xml, "<m:oMath ", fixed = TRUE)
  expect_no_match(xml, "oMathPara", fixed = TRUE)
  # inline style: limits of the sum on the side
  expect_match(xml, "<m:limLoc m:val=\"subSup\"/>", fixed = TRUE)
  expect_match(xml, "Euler: ", fixed = TRUE)
})

test_that("an equation alone in its paragraph is a math paragraph", {
  par <- fpar(
    run_equation("\\sum_{i=1}^{n} x_i"),
    fp_p = fp_par(text.align = "right")
  )
  xml <- to_wml(par)
  expect_match(xml, "<m:oMathPara ", fixed = TRUE)
  expect_match(xml, "<m:jc m:val=\"right\"/>", fixed = TRUE)
  # display style: limits of the sum above and below
  expect_match(xml, "<m:limLoc m:val=\"undOvr\"/>", fixed = TRUE)
  # the alignment of a math paragraph follows the one of its paragraph
  for (align in c("left", "center", "justify")) {
    xml <- to_wml(fpar(run_equation("x"), fp_p = fp_par(text.align = align)))
    jc <- c(left = "left", center = "center", justify = "left")[[align]]
    expect_match(xml, sprintf("<m:jc m:val=\"%s\"/>", jc), fixed = TRUE)
  }
})

test_that("only the size and the color of the run properties are used", {
  prop <- fp_text_lite(font.size = 14, color = "red", font.family = "Arial")
  xml <- to_wml(fpar(run_equation("x", prop = prop)))
  expect_match(xml, "<w:sz w:val=\"28\"/>", fixed = TRUE)
  expect_match(xml, "<w:color w:val=\"FF0000\"/>", fixed = TRUE)
  expect_no_match(xml, "Arial", fixed = TRUE)
  # without properties: those of the document
  xml <- to_wml(fpar(run_equation("x")))
  expect_no_match(xml, "<w:rPr>", fixed = TRUE)
})

test_that("equations are written in a Word document", {
  doc <- read_docx()
  doc <- body_add_fpar(doc, fpar(run_equation("\\frac{a}{b}")))
  doc <- body_add_fpar(doc, fpar("in a sentence: ", run_equation("x^2")))
  path <- print(doc, target = tempfile(fileext = ".docx"))
  dir <- unpack_folder(path, tempfile())
  body <- xml2::read_xml(file.path(dir, "word", "document.xml"))
  ns <- c(
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
  )
  expect_length(xml2::xml_find_all(body, "//w:p/m:oMathPara/m:oMath", ns), 1)
  expect_length(xml2::xml_find_all(body, "//w:p/m:oMath", ns), 1)
})

test_that("PowerPoint: an equation in a sentence is an inline a14:m equation", {
  xml <- to_pml(fpar("Euler: ", run_equation("\\sum_{i=1}^{n} x_i"), "."))
  expect_match(xml, "Requires=\"a14\"><a14:m><m:oMath ", fixed = TRUE)
  expect_no_match(xml, "oMathPara", fixed = TRUE)
  expect_match(xml, "<m:limLoc m:val=\"subSup\"/>", fixed = TRUE)
})

test_that("PowerPoint: an equation alone in its paragraph is a math paragraph", {
  par <- fpar(
    run_equation("\\sum_{i=1}^{n} x_i"),
    fp_p = fp_par(text.align = "center")
  )
  xml <- to_pml(par)
  # the paragraph properties are kept: Excel empties a paragraph without them
  expect_match(xml, "^<a:p><a:pPr ")
  expect_match(xml, "<a14:m><m:oMathPara ", fixed = TRUE)
  expect_match(xml, "<m:jc m:val=\"center\"/>", fixed = TRUE)
  expect_match(xml, "<m:limLoc m:val=\"undOvr\"/>", fixed = TRUE)
})

test_that("PowerPoint: size, color and transparency of the equation", {
  prop <- fp_text_lite(font.size = 20, color = "#FF000080")
  xml <- to_pml(fpar(run_equation("x", prop = prop)))
  expect_match(xml, "<a:rPr sz=\"2000\">", fixed = TRUE)
  expect_match(xml, "<a:srgbClr val=\"FF0000\">\\s*<a:alpha val=\"50196\"/>")
  xml <- to_pml(fpar(run_equation("x")))
  expect_no_match(xml, "<a:rPr", fixed = TRUE)
})

test_that("equations are written in a PowerPoint presentation", {
  pres <- read_pptx()
  pres <- add_slide(pres, layout = "Title and Content")
  pres <- ph_with(
    pres,
    block_list(
      fpar(run_equation("\\frac{a}{b}")),
      fpar("in a sentence: ", run_equation("x^2"))
    ),
    location = ph_location_type("body")
  )
  path <- print(pres, target = tempfile(fileext = ".pptx"))
  dir <- unpack_folder(path, tempfile())
  slide <- xml2::read_xml(file.path(dir, "ppt", "slides", "slide1.xml"))
  ns <- c(
    a14 = "http://schemas.microsoft.com/office/drawing/2010/main",
    m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
  )
  expect_length(xml2::xml_find_all(slide, "//a14:m/m:oMathPara/m:oMath", ns), 1)
  expect_length(xml2::xml_find_all(slide, "//a14:m/m:oMath", ns), 1)
  expect_s3_class(read_pptx(path), "rpptx")
})

xlsx_drawing_xml <- function(x) {
  path <- print(x, target = tempfile(fileext = ".xlsx"))
  dir <- unpack_folder(path, tempfile())
  xml2::read_xml(file.path(dir, "xl", "drawings", "drawing1.xml"))
}
xlsx_ns <- c(
  xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  a = "http://schemas.openxmlformats.org/drawingml/2006/main",
  a14 = "http://schemas.microsoft.com/office/drawing/2010/main",
  m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
)

test_that("Excel: equations are written in a text box", {
  x <- read_xlsx()
  x <- sheet_add_drawing(
    x,
    value = block_list(
      fpar(run_equation("\\frac{a}{b}")),
      fpar("in a sentence: ", run_equation("x^2"))
    ),
    sheet = "Feuil1",
    anchor = "B2",
    width = 3,
    height = 1
  )
  doc <- xlsx_drawing_xml(x)
  sp <- xml2::xml_find_all(doc, "//xdr:oneCellAnchor/xdr:sp", xlsx_ns)
  expect_length(sp, 1)
  expect_identical(
    xml2::xml_attr(xml2::xml_find_first(sp, ".//xdr:cNvSpPr", xlsx_ns), "txBox"),
    "1"
  )
  # paragraphs with their properties: Excel empties a paragraph without them
  expect_length(xml2::xml_find_all(sp, ".//a:p/a:pPr", xlsx_ns), 2)
  expect_length(xml2::xml_find_all(sp, ".//a14:m/m:oMathPara/m:oMath", xlsx_ns), 1)
  expect_length(xml2::xml_find_all(sp, ".//a14:m/m:oMath", xlsx_ns), 1)
  ext <- xml2::xml_find_first(doc, "//xdr:oneCellAnchor/xdr:ext", xlsx_ns)
  expect_identical(xml2::xml_attr(ext, "cx"), "2743200")
})

test_that("Excel: the text box of an equation gets the size of the equation", {
  tex <- "x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}"
  x <- read_xlsx()
  x <- sheet_add_drawing(x, value = fpar(run_equation(tex)), sheet = "Feuil1")
  doc <- xlsx_drawing_xml(x)
  ext <- xml2::xml_find_first(doc, "//xdr:absoluteAnchor/xdr:ext", xlsx_ns)
  s <- equatags::equation_size(tex, size = 11, display = TRUE)
  expect_equal(
    as.numeric(xml2::xml_attr(ext, "cx")) / 914400,
    s$width + 0.2,
    tolerance = 1e-5
  )
  expect_equal(
    as.numeric(xml2::xml_attr(ext, "cy")) / 914400,
    (s$height + s$depth) * 1.2 + 0.1,
    tolerance = 1e-5
  )
})

test_that("Excel: text boxes only contain paragraphs, cells no equations", {
  x <- read_xlsx()
  expect_error(
    sheet_add_drawing(x, value = block_list(fpar("a"), block_toc()), sheet = "Feuil1"),
    "paragraphs made with"
  )
  expect_error(
    sheet_write_data(x, fpar(run_equation("x")), sheet = "Feuil1"),
    "can not be written in the cells"
  )
})
