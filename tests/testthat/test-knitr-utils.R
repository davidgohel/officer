test_that("get_default_pandoc_data_file works as expected", {
  skip_if_not_installed("knitr")
  skip_if_not_installed("rmarkdown")
  skip_if_not(rmarkdown::pandoc_available())

  pptx_pandoc_template <- get_default_pandoc_data_file(
    format = "pptx"
  )
  expect_true(file.exists(pptx_pandoc_template))
  layout_summary_data <- layout_summary(read_pptx(pptx_pandoc_template))
  expect_gt(nrow(layout_summary_data), 0L)
  expect_contains(layout_summary_data$layout, "Title and Content")

  docx_pandoc_template <- get_default_pandoc_data_file(
    format = "docx"
  )
  expect_true(file.exists(docx_pandoc_template))
  doc <- read_docx(docx_pandoc_template)
  styles <- styles_info(doc)
  expect_contains(styles$style_name, "Normal")
  expect_contains(styles$style_name, "Author")
})

test_that("get_reference_value works as expected for docx", {
  skip_if_not_installed("knitr")
  skip_if_not_installed("rmarkdown")
  skip_if_not(rmarkdown::pandoc_available())

  tmp_file <- tempfile(fileext = ".Rmd")
  template_docx <- system.file(package = "officer", "template/template.docx")
  cat(
    "
```{r, include=FALSE}
library(officer)
knitr::knit_hooks$set(dummy = function(before, options, envir) {
  if (!before) {
    template_docx <- get_reference_value(format = \"docx\")
    z <- read_docx(template_docx)
    styles_data <- styles_info(z, type = \"paragraph\")[c(\"style_name\")]

    return(paste(\"```{=openxml}\", to_wml(block_table(styles_data, header = FALSE)), \"```\", sep = \"\n\"))
  }
})
```

```{r, dummy=TRUE}
```

",
    file = tmp_file
  )

  out_file <- tempfile(fileext = ".docx")

  rmarkdown::render(
    input = tmp_file,
    output_format = rmarkdown::word_document(reference_docx = template_docx),
    output_file = out_file,
    quiet = TRUE
  )

  z <- read_docx(out_file)
  content_summary <- docx_summary(z)

  styles <- styles_info(read_docx(), type = "paragraph")

  expect_equal(content_summary$text, styles$style_name)
})

test_that("get_reference_value works as expected for pptx", {
  skip_if_not_installed("knitr")
  skip_if_not_installed("rmarkdown")
  skip_if_not(rmarkdown::pandoc_available())

  pptx_pandoc_template <- get_default_pandoc_data_file(
    format = "pptx"
  )
  template_pptx <- tempfile(fileext = ".pptx")
  file.copy(pptx_pandoc_template, template_pptx)

  str <- "

```{r, include=FALSE}
library(officer)
knitr::knit_hooks$set(dummy = function(before, options, envir) {
  if (!before) {
    template_docx <- get_reference_value(format = \"pptx\")
    saveRDS(template_docx, \"~/dump.RDS\")
    return(\"\")
  }
})
```

```{r, dummy=TRUE}
```

"

  tmp_file <- tempfile(fileext = ".Rmd")
  cat(str, file = tmp_file)

  out_file <- tempfile(fileext = ".pptx")

  rmarkdown::render(
    input = tmp_file,
    output_format = rmarkdown::powerpoint_presentation(
      reference_doc = template_pptx
    ),
    output_file = out_file,
    quiet = TRUE
  )

  expect_equal(basename(readRDS("~/dump.RDS")), basename(template_pptx))
  unlink("~/dump.RDS")
})

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
