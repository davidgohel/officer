test_that("defaul template", {
  x <- read_pptx()
  expect_equal(length( x ), 0)
  expect_true(file.exists(x$package_dir))
})

test_that("console printing", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  expect_output(print(x), "pptx document with 1 ")
})

test_that("check extention and print document", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  print(x, target = "print.pptx")
  expect_true( file.exists("print.pptx") )

  expect_error(print(x, target = "print.pptxxx"))
})


test_that("check template", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  print(x, target = "template.pptx")

  expect_silent(x <- read_pptx(path = "template.pptx"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world again", location = ph_location_type(type = "body"))
  print(x, target = "example.pptx")

  expect_equal(length(x), 2)
})



test_that("slide remove", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 1", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 2", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 3", location = ph_location_type(type = "body"))
  print(x, target = "template.pptx")

  x <- read_pptx(path = "template.pptx")
  x <- remove_slide(x = x)
  expect_equal(length(x), 2)
  x <- remove_slide(x = x, index = 1)
  expect_equal(length(x), 1)

  sm <- slide_summary(x)
  expect_equal(sm[1,]$text, "Hello world 2")
})



test_that("ph remove", {

  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 1", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 2", location = ph_location_type(type = "body"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 3", location = ph_location_type(type = "body"))
  print(x, target = "template.pptx")

  x <- read_pptx(path = "template.pptx")
  x <- ph_remove(x = x, type = "body")

  expect_equal(nrow(slide_summary(x)), 0)
})

test_that("ph remove specific node", {

  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world 1.1", location = ph_location_type(type = "body"))
  x <- ph_with(x, "Hello world 1.2", location = ph_location_type(type = "body"))
  x <- ph_with(x, "Hello world 1.3", location = ph_location_type(type = "body"))


  ## remove just the second ph (with hello world 1.2)
  id_to_remove <- slide_summary(x)$id[slide_summary(x)$text == "Hello world 1.2"]
  x <- ph_remove(x = x, id = id_to_remove)

  expect_equal(
    slide_summary(x)$text,
    c("Hello world 1.1","Hello world 1.3")
  )
  expect_equal(nrow(slide_summary(x)), 2)


})



test_that("cursor is incremented as expected", {
  x <- read_pptx()
  for(i in 1:11){
    x <- add_slide(x, "Title Slide", "Office Theme")
    x <- ph_with(x, i, location = ph_location_type(type = "ctrTitle"))
  }
  expect_equal(nrow( slide_summary(x, 11) ), 1 )
  expect_equal(x$slide$get_slide(11)$name(), "slide11.xml" )

})

test_that("annotate base template", {
  expect_s3_class(try(annotate_base(), silent = TRUE), "rpptx")
})

test_that("no master do not generate an error", {
  x <- read_pptx("docs_dir/no_master.pptx")
  x <- add_slide(x, layout = "Page One", master = "Office Theme")
  x <- try(ph_with(x, i, value = "graphic title",
               location = ph_location_type(type="body")), silent = TRUE)
  expect_s3_class(x, "rpptx")
})



unlink("*.pptx")
unlink("*.emf")

