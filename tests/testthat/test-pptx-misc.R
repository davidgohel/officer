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
  skip_if_not(has_zip())
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with_text(x, type = "body", str = "Hello world")
  print(x, target = "print.pptx")
  expect_true( file.exists("print.pptx") )

  expect_error(print(x, target = "print.pptxxx"))
})


test_that("check template", {
  skip_if_not(has_zip())
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
  skip_if_not(has_zip())
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
  expect_equal(sm$text, "Hello world 2")
})



test_that("ph remove", {
  skip_if_not(has_zip())
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




unlink("*.pptx")
unlink("*.emf")

