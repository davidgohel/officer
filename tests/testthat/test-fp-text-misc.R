context("fp for text - misc")

source("utils.R")

test_that("fp_text - print", {
  fp <- fp_text(font.size = 10)
  expect_output(print(fp))
})

test_that("fp_text - update", {

  fp <- fp_text(font.size = 10)
  expect_equal(fp_sign( fp ), "b219bb0bdd7045575978f22781d0d77a" )

  fp <- update(fp, font.size = 20)
  expect_equal(fp$font.size, 20)
  fp <- update(fp, color = "red")
  expect_equal(fp$color, "red")
  fp <- update(fp, font.family = "Time New Roman")
  expect_equal(fp$font.family, "Time New Roman")
  fp <- update(fp, vertical.align = "superscript")
  expect_equal(fp$vertical.align, "superscript")
  fp <- update(fp, shading.color = "yellow")
  expect_equal(fp$shading.color, "yellow")
})




