img.file <- file.path(R.home("doc"), "html", "logo.jpg")
fpt_blue_bold <- fp_text(color = "#006699", bold = TRUE)
fpt_red_italic <- fp_text(color = "#C32900", italic = TRUE)
bl <- block_list(
  fpar(ftext("hello world", fpt_blue_bold)),
  fpar(
    ftext("hello", fpt_blue_bold),
    " ",
    ftext("world", fpt_red_italic)
  ),
  fpar(
    ftext("hello world", fpt_red_italic),
    external_img(
      src = img.file,
      height = 1.06,
      width = 1.39
    )
  )
)

anyplot <- plot_instr(code = {
  col <- c(
    "#440154FF",
    "#443A83FF",
    "#31688EFF",
    "#21908CFF",
    "#35B779FF",
    "#8FD744FF",
    "#FDE725FF"
  )
  barplot(1:7, col = col, yaxt = "n")
})

test_that("rtf_add works with text, paragraphs, and plots (ggplot2 too)", {
  def_text <- fp_text_lite(color = "#006943", bold = TRUE)
  center_par <- fp_par(text.align = "left", padding = 1, line_spacing = 1.3)

  np <- fp_par(line_spacing = 1.4, padding = 3, )
  fpt_def <- fp_text(
    font.size = 11,
    italic = TRUE,
    bold = TRUE,
    underline = TRUE
  )

  doc <- rtf_doc(normal_par = np, normal_chunk = fpt_def)

  expect_identical(doc$normal_par, np)
  expect_identical(doc$normal_chunk, fpt_def)
  expect_identical(doc$content, list())

  doc <- rtf_add(
    x = doc,
    value = fpar(
      ftext("how are you?", prop = def_text),
      fp_p = fp_par(text.align = "center")
    )
  )

  expect_identical(
    doc$content[[1]]$chunks[[1]],
    ftext("how are you?", prop = def_text)
  )
  expect_identical(doc$content[[1]]$fp_p, fp_par(text.align = "center"))

  a_paragraph <- fpar(
    ftext("Here is a date: ", prop = def_text),
    run_word_field(field = "Date \\@ \"MMMM d yyyy\""),
    fp_p = center_par
  )
  doc <- rtf_add(
    x = doc,
    value = block_list(
      a_paragraph,
      a_paragraph,
      a_paragraph
    )
  )

  expect_identical(doc$content[[4]]$chunks, a_paragraph$chunks)

  if (require("ggplot2")) {
    gg <- gg_plot <- ggplot(data = iris) +
      geom_point(mapping = aes(Sepal.Length, Petal.Length))
    doc <- rtf_add(doc, gg, width = 3, height = 4, ppr = center_par)

    expect_true(grepl("\\.png", doc$content[[5]]$chunks[[1]]))
    expect_identical(
      attr(doc$content[[5]]$chunks[[1]], "dims"),
      list(width = 3, height = 4)
    )
  }
  anyplot <- plot_instr(code = {
    barplot(1:5, col = 2:6)
  })

  doc <- rtf_add(doc, anyplot, width = 5, height = 4, ppr = center_par)
  expect_true(grepl("\\.png", doc$content[[6]]$chunks[[1]]))
  expect_identical(
    attr(doc$content[[6]]$chunks[[1]], "dims"),
    list(width = 5, height = 4)
  )

  expect_s3_class(doc, "rtf")

  expect_identical(
    capture.output(print.rtf(doc)),
    "rtf document with 6 element(s)"
  )

  bl <- block_list(
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

  expect_silent({
    doc <- rtf_add(doc, bl)
  })

  ps <- prop_section(
    page_size = page_size(orient = "landscape"),
    page_margins = page_mar(top = 2),
    type = "continuous"
  )
  bs <- block_section(ps)

  expect_silent({
    doc <- rtf_add(doc, bs)
  })
  expect_silent({
    doc <- rtf_add(doc, "a character")
  })
  expect_silent({
    doc <- rtf_add(doc, factor("a factor"))
  })
  expect_silent({
    doc <- rtf_add(doc, 1.1)
  })

  outfile <- print(doc, target = tempfile(fileext = ".rtf"))
  expect_true(file.exists(outfile))
})
