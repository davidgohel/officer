magick_dims <- function(path) {
  dims <- attr(magick::image_data(magick::image_read(path)), "dim")[-1]
  c(width = dims[1], height = dims[2])
}

gradient <- function(width = 64, height = 48) {
  magick::image_blank(width, height, pseudo_image = "gradient:red-blue")
}

test_that("raster dimensions are returned", {
  r <- as.raster(matrix(0:1, nrow = 2, ncol = 3))
  expect_equal(image_dims(r), c(width = 3, height = 2))

  native <- structure(
    matrix(0L, nrow = 2, ncol = 3),
    class = "nativeRaster",
    channels = 4L
  )
  expect_equal(image_dims(native), c(width = 3, height = 2))
})

test_that("header dimensions are the ones read by magick", {
  skip_if_not_installed("magick")
  images <- list(
    png = magick::image_convert(gradient(), format = "png"),
    png_gray = magick::image_convert(
      gradient(),
      format = "png",
      colorspace = "gray"
    ),
    jpeg = magick::image_convert(gradient(), format = "jpeg"),
    jpeg_gray = magick::image_convert(
      gradient(),
      format = "jpeg",
      colorspace = "gray"
    ),
    jpeg_cmyk = magick::image_convert(
      gradient(),
      format = "jpeg",
      colorspace = "cmyk"
    ),
    jpeg_progressive = magick::image_convert(
      gradient(),
      format = "jpeg",
      interlace = "Plane"
    ),
    gif = magick::image_convert(gradient(), format = "gif")
  )
  for (kind in names(images)) {
    path <- withr::local_tempfile()
    magick::image_write(images[[kind]], path = path)
    expect_equal(image_dims_from_header(path), magick_dims(path), label = kind)
  }
})

test_that("a progressive JPEG has its own frame header", {
  skip_if_not_installed("magick")
  path <- withr::local_tempfile()
  image <- magick::image_convert(
    gradient(),
    format = "jpeg",
    interlace = "Plane"
  )
  magick::image_write(image, path = path)
  bytes <- readBin(path, "raw", n = file.size(path))
  sof2 <- which(bytes[-length(bytes)] == as.raw(0xff) & bytes[-1] == as.raw(0xc2))
  expect_gt(length(sof2), 0L)
  expect_equal(image_dims_from_header(path), c(width = 64, height = 48))
})

test_that("an animated GIF gives the size of its first frame", {
  skip_if_not_installed("magick")
  path <- withr::local_tempfile(fileext = ".gif")
  animation <- magick::image_animate(
    c(gradient(), magick::image_flip(gradient())),
    fps = 5
  )
  magick::image_write(animation, path = path)
  expect_equal(image_dims_from_header(path), magick_dims(path))
})

test_that("metadata before the JPEG frame header are skipped", {
  skip_if_not_installed("magick")
  path <- withr::local_tempfile()
  magick::image_write(gradient(), path = path, format = "jpeg")
  bytes <- readBin(path, "raw", n = file.size(path))
  # APP1 segments of the maximum size, as large EXIF data
  segment <- c(as.raw(c(0xff, 0xe1, 0xff, 0xff)), raw(65533))
  writeBin(c(bytes[1:2], rep(segment, 3), bytes[-(1:2)]), path)
  expect_equal(image_dims_from_header(path), magick_dims(path))
})

test_that("unknown and malformed files are not read from their header", {
  path <- withr::local_tempfile()
  writeLines("this is not an image, not at all", path)
  expect_null(image_dims_from_header(path))

  writeBin(as.raw(c(0xff, 0xd8, 0xff, 0xe0, 0x00)), path)
  expect_null(image_dims_from_header(path))

  # a JPEG without frame header
  writeBin(as.raw(c(0xff, 0xd8, 0xff, 0xd9, 0, 0, 0, 0, 0, 0)), path)
  expect_null(image_dims_from_header(path))

  # a PNG without its IHDR chunk
  writeBin(
    c(as.raw(c(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a)), raw(20)),
    path
  )
  expect_null(image_dims_from_header(path))
})

test_that("other formats are read with magick", {
  skip_if_not_installed("magick")
  path <- withr::local_tempfile(fileext = ".tiff")
  magick::image_write(gradient(), path = path, format = "tiff")
  expect_null(image_dims_from_header(path))
  expect_equal(image_dims(path), c(width = 64, height = 48))
})

test_that("SVG files are read with magick", {
  skip_if_not_installed("magick")
  skip_if_not_installed("rsvg")
  path <- withr::local_tempfile(fileext = ".svg")
  writeLines(
    '<svg xmlns="http://www.w3.org/2000/svg" width="30" height="20"></svg>',
    path
  )
  dims <- attr(magick::image_data(magick::image_read_svg(path)), "dim")[-1]
  expect_equal(image_dims(path), c(width = dims[1], height = dims[2]))
})

test_that("guess_size uses the dimensions of the files", {
  skip_if_not_installed("magick")
  png <- withr::local_tempfile(fileext = ".png")
  jpeg <- withr::local_tempfile(fileext = ".jpeg")
  magick::image_write(gradient(72, 144), path = png, format = "png")
  magick::image_write(gradient(144, 72), path = jpeg, format = "jpeg")

  img <- external_img(c(png, jpeg), guess_size = TRUE)
  expect_equal(dim(img), data.frame(width = c(1, 2), height = c(2, 1)))

  floating <- floating_external_img(png, guess_size = TRUE)
  expect_equal(dim(floating)$width, 1)
  expect_equal(dim(floating)$height, 2)
})
