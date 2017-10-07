context("relationship validity")

source("utils.R")

check_all_types <- function(dat){
  expected_names <- c("id", "int_id", "type", "target", "target_mode", "ext_src")

  expect_is(dat, "data.frame" )
  expect_equal(names(dat), expected_names )

  expect_is(dat$id, "character" )
  expect_is(dat$int_id, "integer" )
  expect_is(dat$type, "character" )
  expect_is(dat$target, "character" )
  expect_is(dat$target_mode, "character" )
  expect_is(dat$ext_src, "character" )
}

test_that("get_data", {

  rel <- officer:::relationship$new()
  dat <- rel$get_data()
  check_all_types(dat)

  rel <- officer:::relationship$new()
  rel$add_img(src = "test.png", root_target = "../media")
  dat <- rel$get_data()
  check_all_types(dat)

  rel <- officer:::relationship$new()
  rel$add_drawing(src = "test.xml", root_target = "../drawings")
  dat <- rel$get_data()
  check_all_types(dat)
})

