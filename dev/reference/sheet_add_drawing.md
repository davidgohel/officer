# Add a drawing to an Excel sheet

Add a graphical element into a sheet of an xlsx workbook. This is a
generic function dispatching on `value`. Supported inputs include:

- [`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md):
  an image file (PNG, JPEG, GIF, ...) copied into the workbook and
  placed on the sheet.

- `gg`: a ggplot object rendered to PNG via
  [`ragg::agg_png()`](https://ragg.r-lib.org/reference/agg_png.html) and
  embedded via the `external_img` path. Extra arguments: `res` (default
  300 ppi), `alt_text` (auto-detected via
  [`ggplot2::get_alt_text()`](https://ggplot2.tidyverse.org/reference/get_alt_text.html)
  when empty), `scale` (same semantics as
  [`ggplot2::ggsave()`](https://ggplot2.tidyverse.org/reference/ggsave.html)).

The drawing can be positioned in three ways, matching Excel's own
"Format Picture" options:

- Fixed position (default): pass `left` / `top` and `width` / `height`
  in inches. The drawing stays put when rows or columns are
  inserted/resized.

- Move and size with cells (Excel's default): pass `anchor = "B2:H20"`
  (a cell range). The drawing spans those two cells and follows them.

- Move but don't size with cells: pass `anchor = "B2"` (a single cell).
  The drawing anchors to that cell, keeping the size set by `width` /
  `height`.

Additional methods for `ms_chart` and `dml` are provided by the
extension packages 'mschart' and 'rvg'.

Use
[`sheet_write_data()`](https://davidgohel.github.io/officer/dev/reference/sheet_write_data.md)
to write data into the sheet before or after adding a drawing.

## Usage

``` r
sheet_add_drawing(x, value, sheet, ...)
```

## Arguments

- x:

  rxlsx object created by
  [`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)

- value:

  object to add (dispatched to the appropriate method)

- sheet:

  sheet name (must already exist, see
  [`add_sheet()`](https://davidgohel.github.io/officer/dev/reference/add_sheet.md))

- ...:

  method-specific arguments. Common across `external_img` and `gg`
  methods: `left` / `top` (fixed-position top-left in inches, default
  `(1, 1)`), `width` / `height` (size in inches), `anchor` (cell
  reference like `"B2"` or `"B2:H20"`, see description), `edit_as` (one
  of `"twoCell"`, `"oneCell"`, `"absolute"`; how Excel treats the
  drawing when rows/columns are resized — only meaningful with a cell
  range `anchor`). The `gg` method also accepts `res` (ppi, default
  300), `alt_text` (auto-detected via
  [`ggplot2::get_alt_text()`](https://ggplot2.tidyverse.org/reference/get_alt_text.html)
  if empty) and `scale` (passed to
  [`ragg::agg_png()`](https://ragg.r-lib.org/reference/agg_png.html)).

## Value

the rxlsx object (invisibly)

## See also

[`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md),
[`add_sheet()`](https://davidgohel.github.io/officer/dev/reference/add_sheet.md),
[`sheet_write_data()`](https://davidgohel.github.io/officer/dev/reference/sheet_write_data.md)

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

if (requireNamespace("ggplot2", quietly = TRUE)) {
  gg <- ggplot2::ggplot(iris, ggplot2::aes(Sepal.Length, Sepal.Width)) +
    ggplot2::geom_point()
  x <- read_xlsx()
  x <- add_sheet(x, label = "plots")
  x <- sheet_add_drawing(x, value = gg, sheet = "plots",
                         width = 4, height = 3)
  print(x, target = tempfile(fileext = ".xlsx"))
}
```
