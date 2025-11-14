# Color scheme of a PowerPoint file

Get the color scheme of a 'PowerPoint' master layout into a data.frame.

## Usage

``` r
color_scheme(x)
```

## Arguments

- x:

  an rpptx object

## See also

Other functions for reading presentation information:
[`annotate_base()`](https://davidgohel.github.io/officer/reference/annotate_base.md),
[`doc_properties()`](https://davidgohel.github.io/officer/reference/doc_properties.md),
[`layout_properties()`](https://davidgohel.github.io/officer/reference/layout_properties.md),
[`layout_summary()`](https://davidgohel.github.io/officer/reference/layout_summary.md),
[`length.rpptx()`](https://davidgohel.github.io/officer/reference/length.rpptx.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/reference/plot_layout_properties.md),
[`slide_size()`](https://davidgohel.github.io/officer/reference/slide_size.md),
[`slide_summary()`](https://davidgohel.github.io/officer/reference/slide_summary.md)

## Examples

``` r
x <- read_pptx()
color_scheme ( x = x )
#>                         name    type   value        theme
#> slideMaster1.xml.1       dk1  sysClr #000000 Office Theme
#> slideMaster1.xml.2       lt1  sysClr #FFFFFF Office Theme
#> slideMaster1.xml.3       dk2 srgbClr #1F497D Office Theme
#> slideMaster1.xml.4       lt2 srgbClr #EEECE1 Office Theme
#> slideMaster1.xml.5   accent1 srgbClr #4F81BD Office Theme
#> slideMaster1.xml.6   accent2 srgbClr #C0504D Office Theme
#> slideMaster1.xml.7   accent3 srgbClr #9BBB59 Office Theme
#> slideMaster1.xml.8   accent4 srgbClr #8064A2 Office Theme
#> slideMaster1.xml.9   accent5 srgbClr #4BACC6 Office Theme
#> slideMaster1.xml.10  accent6 srgbClr #F79646 Office Theme
#> slideMaster1.xml.11    hlink srgbClr #0000FF Office Theme
#> slideMaster1.xml.12 folHlink srgbClr #800080 Office Theme
```
