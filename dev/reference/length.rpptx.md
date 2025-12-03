# Number of slides

Function `length` will return the number of slides.

## Usage

``` r
# S3 method for class 'rpptx'
length(x)
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
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md),
[`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)

## Examples

``` r
my_pres <- read_pptx()
my_pres <- add_slide(my_pres, "Title and Content")
my_pres <- add_slide(my_pres, "Title and Content")
length(my_pres)
#> [1] 2
```
