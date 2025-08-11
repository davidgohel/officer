test_that("settings works", {
  x <- read_docx()
  x <- docx_set_settings(
    x = x,
    zoom = 1,
    default_tab_stop = .5,
    hyphenation_zone = .25,
    decimal_symbol = ".",
    list_separator = ";",
    compatibility_mode = "15",
    even_and_odd_headers = TRUE,
    auto_hyphenation = FALSE
  )
  file <- print(x, target = tempfile(fileext = ".docx"))
  x <- read_docx(path = file)
  expect_equal(x$settings$zoom, 1)
  expect_equal(x$settings$list_separator, ";")
  expect_true(x$settings$even_and_odd_headers)
})
