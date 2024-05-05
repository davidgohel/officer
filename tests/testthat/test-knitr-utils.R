test_that("opts_current_table works as expected", {
  skip_if_not_installed("knitr")
  skip_if_not_installed("rmarkdown")
  skip_if_not(rmarkdown::pandoc_available())

  new_dir <- tempfile()
  dir.create(new_dir, showWarnings = FALSE)
  rmd_fp <- file.path(new_dir, "knitr-utils.Rmd")
  generated_rds <- tempfile(fileext = ".RDS")
  file.copy("docs_dir/knitr-utils.Rmd", to = rmd_fp)
  rmarkdown::render(
    input = rmd_fp,
    envir = parent.frame(),
    params = list(
      rds_path = generated_rds
    ),
    quiet = TRUE
  )
  x <- readRDS(generated_rds)
  expect_equal(x$cap.pre, "Tableau ")
  expect_equal(x$cap.style, "Normal")
  expect_equal(x$cap.sep, " : ")
  expect_equal(x$alt.title, "alt title")
  expect_equal(x$alt.description, "alt description")
  expect_equal(x$cap, "Coucou")
})
