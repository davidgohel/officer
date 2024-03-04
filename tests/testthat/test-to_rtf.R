test_that("to_rtf works with default strings and ftext", {
  str <- "Default string"
  expect_equal(to_rtf(str), str)

  properties2 <- fp_text(bold = TRUE, shading.color = "yellow")
  ft <- ftext("Some text", properties2)
  expect_equal(to_rtf(ft), "%font:Arial%\\b\\fs20%ftcolor:black% %ftshading:yellow%Some text\\highlight0")
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
  # expect_true(grepl("logo\\.jpg", to_rtf(fpar_1)))
})
