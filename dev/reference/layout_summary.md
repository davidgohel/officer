# Presentation layouts summary

Get information about slide layouts and master layouts into a
data.frame. This function returns a data.frame containing all layout and
master names.

## Usage

``` r
layout_summary(x)
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
[`length.rpptx()`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md),
[`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)

## Examples

``` r
my_pres <- read_pptx()
layout_summary ( x = my_pres )
#>              layout       master
#> 1       Title Slide Office Theme
#> 2 Title and Content Office Theme
#> 3    Section Header Office Theme
#> 4       Two Content Office Theme
#> 5        Comparison Office Theme
#> 6        Title Only Office Theme
#> 7             Blank Office Theme
```
