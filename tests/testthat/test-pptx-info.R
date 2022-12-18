test_that("layout summary", {
  x <- read_pptx()
  laysum <- layout_summary(x)
  expect_is( laysum, "data.frame" )
  expect_true( all( c("layout", "master") %in% names(laysum)) )
  expect_is( laysum$layout, "character" )
  expect_is( laysum$master, "character" )
})

test_that("layout properties", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  x <- ph_with(x, "my title", location = ph_location_type(type = "title"))

  laypr <- layout_properties(x, layout = "Title and Content", master = "Office Theme")

  expect_is( laypr, "data.frame" )
  expect_true( all( c("master_name", "name", "type", "offx", "offy", "cx", "cy", "rotation") %in% names(laypr)) )
  expect_is( laypr$master_name, "character" )
  expect_is( laypr$name, "character" )
  expect_is( laypr$type, "character" )
  expect_is( laypr$offx, "numeric" )
  expect_is( laypr$offy, "numeric" )
  expect_is( laypr$cx, "numeric" )
  expect_is( laypr$cy, "numeric" )
  expect_is( laypr$rotation, "numeric" )
})

save_png <- function(code, width = 700, height = 700) {
  path <- tempfile(fileext = ".png")
  png(path, width = width, height = height, res = 150)
  on.exit(dev.off())
  code

  path
}
expect_snapshot_plot <- function(name, code, tolerance = 0.001) {
  skip_on_os("windows")
  name <- paste0(name, ".png")
  announce_snapshot_file(name = name)
  path <- save_png(code)
  expect_snapshot_file(path, name)
}

test_that("plot layout properties", {
  testthat::local_edition(x = 3L)
  x <- read_pptx()
  expect_snapshot_plot("plot-title-layout", code = {
    plot_layout_properties( x = x, layout = "Title Slide",
                            master = "Office Theme" )
  })
  expect_snapshot_plot("plot-twocontent-layout", code = {
    plot_layout_properties( x = x, layout = "Two Content",
                            master = "Office Theme" )
  })
  expect_snapshot_plot("plot-twocontent-layout-nolabel", code = {
    plot_layout_properties( x = x, layout = "Two Content",
                            master = "Office Theme",
                            labels = FALSE)
  })

})

test_that("slide summary", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content", "Office Theme")
  x <- ph_with(x, "Hello world", location = ph_location_type(type = "body"))
  x <- ph_with(x, "my title", location = ph_location_type(type = "title"))

  sm <- slide_summary(x)

  expect_is( sm, "data.frame" )
  expect_equal( nrow(sm), 2 )
  expect_true( all( c("id", "type", "offx", "offy", "cx", "cy") %in% names(sm)) )
  expect_is( sm$id, "character" )
  expect_is( sm$type, "character" )
  expect_true( is.double(sm$offx) )
  expect_true( is.double(sm$offy) )
  expect_true( is.double(sm$cx) )
  expect_true( is.double(sm$cy) )
})

test_that("color scheme", {
  x <- read_pptx()
  cs <- color_scheme(x)

  expect_is( cs, "data.frame" )
  expect_equal( ncol(cs), 4 )
  expect_true( all( c("name", "type", "value", "theme") %in% names(cs)) )
  expect_is( cs$name, "character" )
  expect_is( cs$type, "character" )
  expect_is( cs$value, "character" )
  expect_is( cs$theme, "character" )
})


