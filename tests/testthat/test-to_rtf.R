test_that("to_rtf works with default strings and ftext", {
  expect_equal(to_rtf(NULL), "")

  str <- "Default string"
  expect_equal(to_rtf(str), str)

  properties2 <- fp_text(bold = TRUE, shading.color = "yellow")
  ft <- ftext("Some text", properties2)
  expect_equal(
    to_rtf(ft),
    "%font:Arial%\\b\\fs20%ftcolor:black% %ftshading:yellow%Some text\\highlight0"
  )
})

test_that("to_rtf works with fpar and external images", {
  img.file <- file.path(R.home("doc"), "html", "logo.jpg")

  bold_face <- shortcuts$fp_bold(font.size = 12)
  bold_redface <- update(bold_face, color = "red")
  fpar_1 <- fpar(
    "Hello World, ",
    ftext("how ", prop = bold_redface),
    external_img(src = img.file, height = 1.06 / 2, width = 1.39 / 2),
    ftext(" you?", prop = bold_face)
  )
  expect_true(grepl("Hello World, ", to_rtf(fpar_1)))
})

test_that("to_rtf works with run_word_field, run_pagebreak, run_columnbreak, and run_linebreak", {
  out <- to_rtf(run_word_field(field = "PAGE  \\* MERGEFORMAT"))

  expect_true(grepl("PAGE", out))

  fp_t <- fp_text(font.size = 12, bold = TRUE)
  an_fpar <- fpar(
    "let's add a break page",
    run_pagebreak(),
    ftext("and blah blah!", fp_t)
  )

  expect_true(grepl(" let's add a break page", to_rtf(an_fpar)))
  expect_true(grepl("and blah blah", to_rtf(an_fpar)))

  expect_true(grepl("column", to_rtf(run_columnbreak())))

  fp_t <- fp_text(font.size = 12, bold = TRUE)
  an_fpar <- fpar(
    "let's add a line break",
    run_linebreak(),
    ftext("and blah blah!", fp_t)
  )

  expect_true(grepl(" let's add a line break", to_rtf(an_fpar)))
  expect_true(grepl("and blah blah", to_rtf(an_fpar)))
})

test_that("to_rtf works with hyperlinks and block_list", {
  ft <- fp_text(font.size = 12, bold = TRUE)
  ft <- hyperlink_ftext(
    href = "https://cran.r-project.org/index.html",
    text = "some text",
    prop = ft
  )

  expect_true(grepl("HYPERLINK", to_rtf(ft)))

  fpt_blue_bold <- fp_text(color = "#006699", bold = TRUE)
  fpt_red_italic <- fp_text(color = "#C32900", italic = TRUE)

  value <- block_list(
    fpar(ftext("hello world\\t", fpt_blue_bold)),
    fpar(
      ftext("hello", fpt_blue_bold),
      " ",
      ftext("world", fpt_red_italic)
    ),
    fpar(
      ftext("hello world", fpt_red_italic)
    )
  )

  expect_true(grepl("C32900", to_rtf(value)))
  expect_true(grepl("006699", to_rtf(value)))
  expect_true(grepl("hello world", to_rtf(value)))
})

test_that("to_rtf works for run_autonum", {
  ra <- run_autonum(
    seq_id = "tab",
    pre_label = "Table ",
    bkm = "anytable",
    tnd = 2,
    tns = " "
  )

  expect_true(grepl("bkmkstart anytable", to_rtf(ra)))
  expect_true(grepl("SEQ tab", to_rtf(ra)))

  rr <- run_reference("a_ref")

  expect_true(grepl("REF a_ref ", to_rtf(rr)))

  ps <- prop_section(
    page_size = page_size(orient = "landscape"),
    page_margins = page_mar(top = 2),
    type = "continuous"
  )
  bs <- block_section(ps)

  expect_true(grepl("lndscpsxn", to_rtf(bs)))
  expect_true(grepl("pghsxn11906", to_rtf(bs)))
  expect_true(grepl("margt2880", to_rtf(bs))) # top = 2

  expect_true(grepl("margb1417", to_rtf(page_mar())))
  expect_true(grepl("colw3600", to_rtf(section_columns())))
  expect_true(grepl("pgwsxn16838", to_rtf(page_size(orient = "landscape"))))
  expect_true(grepl("lndscpsxn", to_rtf(page_size(orient = "landscape"))))
})

test_that("tcpr_rtf is retrieved by format", {
  obj <- fp_cell(margin = 1)
  fp_c <- update(obj, margin.bottom = 5)

  expect_true(grepl("clvertalc", format(fp_c, "rtf")))
})

test_that("to_rtf with fp_par_lite", {
  x <- fpar("hello", fp_p = fp_par_lite())
  expect_equal(to_rtf(x), "{hello\\par}")
  x <- fpar("hello", fp_p = fp_par_lite(text.align = "center"))
  expect_equal(to_rtf(x), "{\\pard\\qc hello\\par}")
})

test_that("border_rtf", {
  x <- fp_border()
  expect_equal(
    border_rtf(x, side = "right"),
    "\\clbrdrr\\brdrs\\brdrw20%ftlinecolor:black%"
  )
  expect_equal(
    border_rtf(x, side = "left"),
    "\\clbrdrl\\brdrs\\brdrw20%ftlinecolor:black%"
  )
  expect_equal(
    border_rtf(x, side = "top"),
    "\\clbrdrt\\brdrs\\brdrw20%ftlinecolor:black%"
  )
  expect_equal(
    border_rtf(x, side = "bottom"),
    "\\clbrdrb\\brdrs\\brdrw20%ftlinecolor:black%"
  )
})

test_that("fp_text_lite and rtf", {
  x <- fp_text_lite()
  expect_equal(format(x, type = "rtf"), "")
  x <- fp_text_lite(vertical.align = "superscript", font.size = 12)
  expect_equal(format(x, type = "rtf"), "\\super\\up6\\fs24 ")
  x <- fp_text_lite(vertical.align = "subscript", font.size = 12)
  expect_equal(format(x, type = "rtf"), "\\sub\\dn6\\fs24 ")
})

test_that("to_rtf.section_columns emits per-column widths (#726)", {
  x <- section_columns(widths = c(2, 3), space = 0.25)
  out <- officer:::to_rtf.section_columns(x)
  expect_equal(
    out,
    "\\cols2\\colsx360\\colno1\\colw2880\\colsr360\\colno2\\colw4320"
  )

  # sep = TRUE adds \linebetcol just after the global \colsx
  x2 <- section_columns(widths = c(2, 3), space = 0.25, sep = TRUE)
  expect_equal(
    officer:::to_rtf.section_columns(x2),
    "\\cols2\\colsx360\\linebetcol\\colno1\\colw2880\\colsr360\\colno2\\colw4320"
  )

  # three columns: two \colsr entries (one per non-last column)
  x3 <- section_columns(widths = c(1, 2, 3), space = 0.5)
  out3 <- officer:::to_rtf.section_columns(x3)
  expect_match(out3, "\\\\colno1\\\\colw1440\\\\colsr720")
  expect_match(out3, "\\\\colno2\\\\colw2880\\\\colsr720")
  expect_match(out3, "\\\\colno3\\\\colw4320(?!\\\\colsr)", perl = TRUE)
})

test_that("section properties precede \\sect in block_section (#726)", {
  bs <- block_section(prop_section(
    section_columns = section_columns(widths = c(2, 3), space = 0.25)
  ))
  out <- to_rtf(bs)
  # \sect must come at the very end, after all section properties
  expect_match(out, "\\\\cols2\\\\colsx360.*\\\\sect$")
  expect_false(grepl("^\\\\sect", out))
})

test_that("run_columnbreak emits a delimiter after \\column (#726)", {
  # The trailing space prevents "\columnRight" from being parsed as a
  # single unknown control word when text follows directly.
  expect_equal(to_rtf(run_columnbreak()), "\\column ")
})
