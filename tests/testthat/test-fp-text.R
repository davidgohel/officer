context("fp_text")

library(utils)
library(xml2)


test_that("fp_text - print", {
  fp <- fp_text(font.size = 10)
  expect_output(print(fp))
})

test_that("fp_text - as.data.frame", {
  fp <- fp_text(font.size = 10, color = "red", bold = TRUE, italic = TRUE, underlined = TRUE, font.family = "Arial", shading.color = "yellow", vertical.align = "superscript")
  expect_is(as.data.frame(fp), class = "data.frame")
})

test_that("fp_text - update", {
  fp <- fp_text(font.size = 10)
  fp <- update(fp, font.size = 20)
  expect_equal(fp$font.size, 20)
  fp <- update(fp, color = "red")
  expect_equal(fp$color, "red")
  fp <- update(fp, bold = TRUE)
  expect_equal(fp$bold, TRUE)
  fp <- update(fp, italic = TRUE)
  expect_equal(fp$italic, TRUE)
  fp <- update(fp, underlined = TRUE)
  expect_equal(fp$underlined, TRUE)
  fp <- update(fp, font.family = "Time New Roman")
  expect_equal(fp$font.family, "Time New Roman")
  fp <- update(fp, vertical.align = "superscript")
  expect_equal(fp$vertical.align, "superscript")
  fp <- update(fp, shading.color = "yellow")
  expect_equal(fp$shading.color, "yellow")
})
