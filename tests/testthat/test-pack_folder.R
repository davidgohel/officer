context("packing and unpacking folders")

test_that("packing and unpacking does not change files", {

  fold_fname <- tempfile()
  docx_fname <- paste(fold_fname, '/', basename(fold_fname), '.docx', sep = "")
  zip_fname <- tempfile(fileext = '.zip')
  system(paste('mkdir', fold_fname))
  system(paste('mkdir', basename(zip_fname)))

  x <- read_docx() %>%
    body_add_par("paragraph 1", style = "Normal") %>%
    cursor_begin()

  print(x, docx_fname) %>% invisible()

  pack_folder(fold_fname, zip_fname)
  unpack_folder(zip_fname, fold_fname)

  y <- read_docx(docx_fname)
  expect_equal(x$doc_obj$get(), y$doc_obj$get())
})
