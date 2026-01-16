# Write a ggplot Object to PNG File

Renders a ggplot object to a PNG file using ragg for high-quality
output.

## Usage

``` r
plot_in_png(
  ggobj = NULL,
  code = NULL,
  width,
  height,
  res = 200,
  units = "in",
  pointsize = 11,
  scaling = 1,
  path = NULL
)
```

## Arguments

- ggobj:

  A ggplot object to render

- width:

  Numeric, width of the output image

- height:

  Numeric, height of the output image

- res:

  Numeric, resolution in DPI (default 200)

- units:

  Character, units for width and height ("in", "cm", "mm", "px")
  (default "in")

- pointsize:

  Integer, The default pointsize of the device in pt

- scaling:

  scaling factor to apply

- path:

  Character, output file path. If NULL, a temporary file is created
  (default NULL)

## Value

Character, the path to the created PNG file

## Examples

``` r
plot_in_png(
  code = {
    barplot(1:10)
  },
  width = 5,
  height = 4,
  res = 72,
  units = "in"
)
#> [1] "/tmp/RtmpMgcMop/file179a10d1782d.png"
```
