# Section columns

The function creates a representation of the columns of a section.

## Usage

``` r
section_columns(widths = c(2.5, 2.5), space = 0.25, sep = FALSE)
```

## Arguments

- widths:

  columns widths in inches. If 3 values, 3 columns will be produced.

- space:

  space in inches between columns.

- sep:

  if TRUE a line is separating columns.

## See also

Other functions for section definition:
[`page_mar()`](https://davidgohel.github.io/officer/dev/reference/page_mar.md),
[`page_size()`](https://davidgohel.github.io/officer/dev/reference/page_size.md),
[`prop_section()`](https://davidgohel.github.io/officer/dev/reference/prop_section.md)

## Examples

``` r
section_columns()
#> $widths
#> [1] 2.5 2.5
#> 
#> $space
#> [1] 0.25
#> 
#> $sep
#> [1] FALSE
#> 
#> attr(,"class")
#> [1] "section_columns"
```
