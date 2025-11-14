# 'Word' page layout

Get page width, page height and margins (in inches). The return values
are those corresponding to the section where the cursor is.

## Usage

``` r
docx_dim(x)
```

## Arguments

- x:

  an `rdocx` object

## See also

Other functions for Word document informations:
[`doc_properties()`](https://davidgohel.github.io/officer/reference/doc_properties.md),
[`docx_bookmarks()`](https://davidgohel.github.io/officer/reference/docx_bookmarks.md),
[`length.rdocx()`](https://davidgohel.github.io/officer/reference/length.rdocx.md),
[`set_doc_properties()`](https://davidgohel.github.io/officer/reference/set_doc_properties.md),
[`styles_info()`](https://davidgohel.github.io/officer/reference/styles_info.md)

## Examples

``` r
docx_dim(read_docx())
#> $page
#>     width    height 
#>  8.263889 11.694444 
#> 
#> $landscape
#> [1] FALSE
#> 
#> $margins
#>       top    bottom      left     right    header    footer 
#> 0.9840278 0.9840278 0.9840278 0.9840278 0.4916667 0.4916667 
#> 
```
