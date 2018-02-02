context("miscellaneous checks for pptx")

test_that("defaul template", {
  x <- read_pptx()
  expect_equal(length( x ), 0)
  expect_true(file.exists(x$package_dir))
})

test_that("console printing", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  expect_output(print(x), "pptx document with 1 ")
})

test_that("check extention and print document", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  print(x, target = "print.pptx")
  expect_true( file.exists("print.pptx") )

  expect_error(print(x, target = "print.pptxxx"))
})


test_that("check template", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  print(x, target = "template.pptx")

  expect_silent(x <- read_pptx(path = "template.pptx"))
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world again")
  print(x, target = "example.pptx")

  expect_equal(length(x), 2)
})



test_that("slide remove", {
  x <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 1") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 2") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 3")
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
  x <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 1") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 2") %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "Hello world 3")
  print(x, target = "template.pptx")

  x <- read_pptx(path = "template.pptx")
  x <- ph_remove(x = x)

  expect_equal(nrow(slide_summary(x)), 0)
})


test_that("cursor is incremented as expected", {
  x <- read_pptx()
  for(i in 1:11){
    x <- add_slide(x, "Title Slide", "Office Theme")
    x <- ph_with_text(x, i, type = "ctrTitle")
  }
  expect_equal(nrow( slide_summary(x, 11) ), 1 )
  expect_equal(x$slide$get_slide(11)$name(), "slide11.xml" )

})

test_that("annotate base template", {
  annotate_base() %>% invisible()
})

unlink("*.pptx")
unlink("*.emf")

