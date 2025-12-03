# Slides width and height

Get the width and height of slides in inches as a named vector.

## Usage

``` r
slide_size(x)
```

## Arguments

- x:

  an rpptx object

## See also

Other functions for reading presentation information:
[`annotate_base()`](https://davidgohel.github.io/officer/dev/reference/annotate_base.md),
[`color_scheme()`](https://davidgohel.github.io/officer/dev/reference/color_scheme.md),
[`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md),
[`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md),
[`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md),
[`length.rpptx()`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)

## Examples

``` r
my_pres <- read_pptx()
my_pres <- add_slide(my_pres,
  layout = "Two Content", master = "Office Theme")
slide_size(my_pres)
#> $width
#> [1] 10
#> 
#> $height
#> [1] 7.5
#> 
```
