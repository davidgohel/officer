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

  x <- rtf_doc()
  x <- rtf_add(x, "Hello World")
  x <- rtf_add(x, anyplot)
  x <- rtf_add(x = x, bl)

  expect_snapshot_doc(x = x, name = "rtf-elements", engine = "testthat")
})

