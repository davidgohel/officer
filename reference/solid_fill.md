# Solid fill DrawingML fragment

Build an `<a:solidFill>` block with an sRGB color and an alpha channel.
Used inside DrawingML shape properties wrappers such as `<a:spPr>`,
`<a:ln>`, `<a:rPr>`, and chartEx-specific wrappers such as `<cx:spPr>`.

This is a low-level helper intended for packages that emit OOXML
fragments (e.g. `mschart`, `rvg`). End users typically rely on
higher-level objects
([`fp_text()`](https://davidgohel.github.io/officer/reference/fp_text.md),
[`fp_border()`](https://davidgohel.github.io/officer/reference/fp_border.md),
[`sp_line()`](https://davidgohel.github.io/officer/reference/sp_line.md))
that embed the `<a:solidFill>` block themselves.

## Usage

``` r
solid_fill(color)
```

## Arguments

- color:

  a color value: a name (`"red"`), a hex string with or without leading
  `#` (`"#FF0000"` or `"FF0000"`), or `"transparent"`. Anything accepted
  by [`grDevices::col2rgb()`](https://rdrr.io/r/grDevices/col2rgb.html)
  works; the alpha channel is honored.

## Value

a character string with the OOXML fragment, namespaced `a:`.

## Examples

``` r
solid_fill("red")
#> [1] "<a:solidFill><a:srgbClr val=\"FF0000\"><a:alpha val=\"100000\"/></a:srgbClr></a:solidFill>"
solid_fill("#3366CC")
#> [1] "<a:solidFill><a:srgbClr val=\"3366CC\"><a:alpha val=\"100000\"/></a:srgbClr></a:solidFill>"
```
