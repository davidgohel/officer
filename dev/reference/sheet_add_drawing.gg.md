# Add a ggplot to an Excel sheet

Render a ggplot object to PNG and embed it into a sheet of an xlsx
workbook via
[`sheet_add_drawing.external_img()`](https://davidgohel.github.io/officer/dev/reference/sheet_add_drawing.external_img.md).
For an editable vector-graphics version, use `rvg::dml(ggobj = ...)` +
`rvg::sheet_add_drawing.dml()` instead.

## Usage

``` r
# S3 method for class 'gg'
sheet_add_drawing(
  x,
  value,
  sheet,
  left = 1,
  top = 1,
  width = 6,
  height = 4,
  res = 300,
  alt_text = "",
  scale = 1,
  ...
)
```

## Arguments

- x:

  rxlsx object created by
  [`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)

- value:

  a ggplot object (class `gg`)

- sheet:

  sheet name (must already exist)

- left, top:

  top-left anchor of the image, in inches

- width, height:

  size of the image, in inches

- res:

  resolution of the png image, in ppi (default 300)

- alt_text:

  alt text for screen readers. If `""` or `NULL`,
  [`ggplot2::get_alt_text()`](https://ggplot2.tidyverse.org/reference/get_alt_text.html)
  is consulted.

- scale:

  multiplicative scaling factor, same as `ggplot2::ggsave(scale=)`

- ...:

  passed to
  [`ragg::agg_png()`](https://ragg.r-lib.org/reference/agg_png.html)

## Value

the rxlsx object (invisibly)

## See also

[`sheet_add_drawing()`](https://davidgohel.github.io/officer/dev/reference/sheet_add_drawing.md),
[`sheet_add_drawing.external_img()`](https://davidgohel.github.io/officer/dev/reference/sheet_add_drawing.external_img.md)
