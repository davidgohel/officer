img.file <- file.path(R.home("doc"), "html", "logo.jpg")
fpt_blue_bold <- fp_text(color = "#006699", bold = TRUE)
fpt_red_italic <- fp_text(color = "#C32900", italic = TRUE)
bl <- block_list(
  fpar(ftext("hello world", fpt_blue_bold)),
  fpar(
    ftext("hello", fpt_blue_bold), " ",
    ftext("world", fpt_red_italic)
  ),
  fpar(
    ftext("hello world", fpt_red_italic),
    external_img(
      src = img.file, height = 1.06, width = 1.39
    )
  )
)

anyplot <- plot_instr(code = {
  col <- c(
    "#440154FF", "#443A83FF", "#31688EFF",
    "#21908CFF", "#35B779FF", "#8FD744FF", "#FDE725FF"
  )
  barplot(1:7, col = col, yaxt = "n")
})

bl <- block_list(
  fpar(ftext("hello world\\t", fpt_blue_bold)),
  fpar(
    ftext("hello", fpt_blue_bold), " ",
    ftext("world", fpt_red_italic)
  ),
  fpar(
    ftext("hello world", fpt_red_italic)
  )
)

ps <- prop_section(
  page_size = page_size(orient = "landscape"),
  page_margins = page_mar(top = 2),
  type = "continuous"
)
bs <- block_section(ps)

run_num <- run_autonum(
  seq_id = "tab", pre_label = "tab. ",
  bkm = "mtcars_table"
)
caption <- block_caption("mtcars table",
                         style = "Normal",
                         autonum = run_num
)
fp_t <- fp_text(font.size = 12, bold = TRUE)
an_fpar <- fpar("let's add a break page", run_pagebreak(), ftext("and blah blah!", fp_t))

test_that("visual testing", {
  doc <- read_docx()
  # add text and a table ----
  doc <- body_add_par(doc, "Hello World")
  doc <- body_add_par(doc, "Hello title", style = "heading 1")
  doc <- body_add_par(doc, "Hello title", style = "heading 2")
  doc <- body_add_table(doc, head(cars))
  doc <- body_add_par(doc, "Hello base plot", style = "heading 2")
  doc <- body_add_plot(doc, anyplot)
  doc <- body_add_par(doc, "Hello fpars", style = "heading 2")
  doc <- body_add_blocks(doc, blocks = bl)
  doc <- body_add(doc, "some char")
  doc <- body_add(doc, 1.1)
  doc <- body_add(doc, factor("a factor"))
  doc <- body_add(doc, fpar(ftext("hello", shortcuts$fp_bold())))
  doc <- body_add(doc, external_img(src = img.file, height = 1.06 / 2, width = 1.39 / 2))
  doc <- body_add(doc, data.frame(mtcars))
  doc <- body_add(doc, bl)
  doc <- body_add(doc, bs)
  doc <- body_add(doc, caption)
  doc <- body_add(doc, block_toc(style = "Table Caption"))
  doc <- body_add(doc, an_fpar)
  doc <- body_add(doc, run_columnbreak())
  if (require("ggplot2")) {
    gg <- gg_plot <- ggplot(data = iris) +
      geom_point(mapping = aes(Sepal.Length, Petal.Length))
    doc <- body_add(doc, gg,
                    width = 3, height = 4
    )
  }
  doc <- body_add(doc, anyplot)

  expect_silent(print(doc, target = "external_file.docx"))

  local_edition(3)
  testthat::skip_if_not_installed("doconv")
  testthat::skip_if_not(doconv::msoffice_available())
  library(doconv)
  expect_snapshot_doc(doc, name = "docx-elements", engine = "testthat")
})

# test_that("body_add visual testing", {
#   local_edition(3)
#   testthat::skip_if_not_installed("doconv")
#   testthat::skip_if_not(doconv::msoffice_available())
#   library(doconv)
#
#   x <- read_docx()
#   # add text and a table ----
#   x <- body_add(x, "Hello World")
#   x <- body_add(x, "Hello title", style = "heading 1")
#   x <- body_add(x, "Hello title", style = "heading 2")
#   x <- body_add(x, head(cars))
#   x <- body_add(x, "Hello base plot", style = "heading 2")
#   x <- body_add(x, anyplot)
#   x <- body_add(x, "Hello fpars", style = "heading 2")
#   x <- body_add(x = x, bl)
#
#   expect_snapshot_doc(x = x, name = "body_add-elements", engine = "testthat")
# })
