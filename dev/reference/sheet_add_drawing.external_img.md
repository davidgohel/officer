# Add an image to an Excel sheet

Add an image file (PNG, JPEG, GIF, ...) to a sheet in an xlsx workbook
created with
[`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md).
The image is copied into `xl/media/` and placed on the sheet via an
absolute anchor (inch-based position and size).

## Usage

``` r
# S3 method for class 'external_img'
sheet_add_drawing(
  x,
  value,
  sheet,
  left = 1,
  top = 1,
  width = NULL,
  height = NULL,
  ...
)
```

## Arguments

- x:

  rxlsx object (created by
  [`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md))

- value:

  an
  [`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md)
  object

- sheet:

  sheet name (must already exist)

- left, top:

  top-left anchor of the image, in inches. Defaults to `(1, 1)`.

- width, height:

  size of the image, in inches. When `NULL` (default) the dimensions
  stored on `value` (from
  [`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md))
  are used.

- ...:

  unused

## Value

the rxlsx object (invisibly)

## See also

[`sheet_add_drawing()`](https://davidgohel.github.io/officer/dev/reference/sheet_add_drawing.md),
[`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md)

## Examples

``` r
img <- system.file("extdata", "example.png", package = "officer")
if (nzchar(img) && file.exists(img)) {
  x <- read_xlsx()
  x <- add_sheet(x, label = "pics")
  x <- sheet_add_drawing(
    x, sheet = "pics",
    value = external_img(img, width = 2, height = 2)
  )
  print(x, target = tempfile(fileext = ".xlsx"))
}
```
