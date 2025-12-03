# Read document properties

Read Word or PowerPoint document properties and get results in a
data.frame.

## Usage

``` r
doc_properties(x)
```

## Arguments

- x:

  an `rdocx` or `rpptx` object

## Value

a data.frame

## See also

Other functions for Word document informations:
[`docx_bookmarks()`](https://davidgohel.github.io/officer/dev/reference/docx_bookmarks.md),
[`docx_dim()`](https://davidgohel.github.io/officer/dev/reference/docx_dim.md),
[`length.rdocx()`](https://davidgohel.github.io/officer/dev/reference/length.rdocx.md),
[`set_doc_properties()`](https://davidgohel.github.io/officer/dev/reference/set_doc_properties.md),
[`styles_info()`](https://davidgohel.github.io/officer/dev/reference/styles_info.md)

Other functions for reading presentation information:
[`annotate_base()`](https://davidgohel.github.io/officer/dev/reference/annotate_base.md),
[`color_scheme()`](https://davidgohel.github.io/officer/dev/reference/color_scheme.md),
[`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md),
[`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md),
[`length.rpptx()`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md),
[`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)

## Examples

``` r
x <- read_docx()
doc_properties(x)
#>               tag                value
#> 1           title                     
#> 2         subject                     
#> 3         creator                     
#> 4        keywords                     
#> 5     description                     
#> 6  lastModifiedBy          David Gohel
#> 7        revision                    9
#> 8         created 2017-02-28T11:18:00Z
#> 9        modified 2020-05-13T09:11:00Z
#> 10       category                     
```
