test_that("layout summary", {
  x <- read_pptx()
  laysum <- layout_summary(x)
  expect_is(laysum, "data.frame")
  expect_true(all(c("layout", "master") %in% names(laysum)))
  expect_is(laysum$layout, "character")
  expect_is(laysum$master, "character")
})


test_that("layout summary - layout order (#596)", {
  file <- test_path("docs_dir", "test-layouts-ordering.pptx")
  x <- read_pptx(file)
  df <- layout_summary(x)
  order_exp <- c(
    "Title Slide", "Title and Content", "Section Header", "Two Content", "Comparison",
    "Title Only", "Blank", "layout_8", "layout_9", "layout_10", "layout_11"
  )
  expect_equal(df$layout, order_exp)
  df <- x$slideLayouts$get_metadata() # used inside layout_summary
  expect_true(all(get_file_index(df$filename) == 1:11))

  file <- test_path("docs_dir", "test-layouts-ordering-3-masters.pptx")
  x <- read_pptx(file)
  df <- layout_summary(x)
  la <- c("Title Slide", "Title and Content", "Section Header", "Two Content", "Comparison", "Title Only", "Blank")
  order_exp <- rep(la, 3)
  expect_equal(df$layout, order_exp)
  order_exp <- rep(paste0("Master_", 1:3), each = length(la))
  expect_equal(df$master, order_exp)
})


test_that("layout properties", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  x <- ph_with(x, "my title", location = ph_location_type(type = "title"))

  laypr <- layout_properties(x, layout = "Title and Content", master = "Office Theme")

  expect_is(laypr, "data.frame")
  expect_true(all(c("master_name", "name", "type", "offx", "offy", "cx", "cy", "rotation") %in% names(laypr)))
  expect_is(laypr$master_name, "character")
  expect_is(laypr$name, "character")
  expect_is(laypr$type, "character")
  expect_is(laypr$offx, "numeric")
  expect_is(laypr$offy, "numeric")
  expect_is(laypr$cx, "numeric")
  expect_is(laypr$cy, "numeric")
  expect_is(laypr$rotation, "numeric")
})


test_that("layout properties - all phs for multiple masters (#597)", {
  file <- test_path("docs_dir/test-three-identical-masters.pptx")
  x <- read_pptx(file)
  lap <- layout_properties(x)

  expect_true(all(table(lap$ph) == 3)) # 3 identical masters => each ph 3 times

  l_df <- split(lap, lap$master_name) # phs sorted by y coords
  is_y_sorted <- vapply(l_df, function(x) {
    all(diff(x$offy) >= 0)
  }, logical(1))
  expect_true(all(is_y_sorted))
})


save_png <- function(code, width = 700, height = 700) {
  path <- tempfile(fileext = ".png")
  png(path, width = width, height = height, res = 150)
  on.exit(dev.off())
  code

  path
}


test_that("plot layout properties", {
  skip_if_not_installed("doconv")
  skip_if_not(doconv::msoffice_available())
  require(doconv)
  local_edition(3L)
  x <- read_pptx()

  png1 <- tempfile(fileext = ".png")
  png(png1, width = 7, height = 6, res = 150, units = "in")
  plot_layout_properties(
    x = x, layout = "Title Slide", master = "Office Theme"
  )
  dev.off()

  png2 <- tempfile(fileext = ".png")
  png(png2, width = 7, height = 6, res = 150, units = "in")
  plot_layout_properties(
    x = x, layout = "Title Slide", master = "Office Theme",
    labels = TRUE, type = FALSE, id = FALSE, title = FALSE
  )
  dev.off()

  png3 <- tempfile(fileext = ".png")
  png(png3, width = 7, height = 6, res = 150, units = "in")
  plot_layout_properties(
    x = x, layout = "Title Slide", master = "Office Theme", legend = TRUE
  )
  dev.off()

  expect_snapshot_doc(name = "plot-titleslide-layout-default", x = png1, engine = "testthat")
  expect_snapshot_doc(name = "plot-titleslide-layout-labels-only", x = png2, engine = "testthat")
  expect_snapshot_doc(name = "plot-titleslide-layout-default-with-legend", x = png3, engine = "testthat")

  # issue #604
  p <- test_path("docs_dir/test-content-order.pptx")
  x <- read_pptx(p)

  png4 <- tempfile(fileext = ".png")
  png(png4, width = 7, height = 6, res = 150, units = "in")
  plot_layout_properties(
    x = x, layout = "Many Contents", master = "Office Theme"
  )
  dev.off()

  png5 <- tempfile(fileext = ".png")
  png(png5, width = 7, height = 6, res = 150, units = "in")
  plot_layout_properties(
    x = x, layout = "Many Contents", master = "Office Theme",
    labels = TRUE, type = FALSE, id = FALSE, title = FALSE
  )
  dev.off()

  expect_snapshot_doc(name = "plot-content-order-default", x = png4, engine = "testthat")
  expect_snapshot_doc(name = "plot-content-order-labels-only", x = png5, engine = "testthat")
})


test_that("slide summary", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  x <- ph_with(x, "my title", location = ph_location_type(type = "title"))

  sm <- slide_summary(x)

  expect_is(sm, "data.frame")
  expect_equal(nrow(sm), 2)
  expect_true(all(c("id", "type", "offx", "offy", "cx", "cy") %in% names(sm)))
  expect_is(sm$id, "character")
  expect_is(sm$type, "character")
  expect_true(is.double(sm$offx))
  expect_true(is.double(sm$offy))
  expect_true(is.double(sm$cx))
  expect_true(is.double(sm$cy))
})



test_that("plot layout properties: layout arg takes numeric index", {
  x <- read_pptx()
  las <- layout_summary(x)
  ii <- as.numeric(rownames(las))

  discarded_plot <- function(x, layout = NULL, master = NULL) { # avoid Rplots.pdf creation
    file <- tempfile(fileext = ".png")
    png(file, width = 7, height = 6, res = 150, units = "in")
      plot_layout_properties(x, layout, master)
    dev.off()
  }

  for (idx in ii) {
    expect_no_error(discarded_plot(x, idx))
  }
  expect_no_error(discarded_plot(x, 1, "Office Theme"))
})


test_that("color scheme", {
  x <- read_pptx()
  cs <- color_scheme(x)

  expect_is(cs, "data.frame")
  expect_equal(ncol(cs), 4)
  expect_true(all(c("name", "type", "value", "theme") %in% names(cs)))
  expect_is(cs$name, "character")
  expect_is(cs$type, "character")
  expect_is(cs$value, "character")
  expect_is(cs$theme, "character")
})
