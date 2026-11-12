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
