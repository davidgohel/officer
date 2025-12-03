# Table conditional formatting

Tables can be conditionally formatted based on few properties as whether
the content is in the first row, last row, first column, or last column,
or whether the rows or columns are to be banded.

## Usage

``` r
table_conditional_formatting(
  first_row = TRUE,
  first_column = FALSE,
  last_row = FALSE,
  last_column = FALSE,
  no_hband = FALSE,
  no_vband = TRUE
)
```

## Arguments

- first_row, last_row:

  apply or remove formatting from the first or last row in the table.

- first_column, last_column:

  apply or remove formatting from the first or last column in the table.

- no_hband, no_vband:

  don't display odd and even rows or columns with alternating shading
  for ease of reading.

## Note

You must define a format for first_row, first_column and other
properties if you need to use them. The format is defined in a docx
template.

## See also

Other functions for table definition:
[`prop_table()`](https://davidgohel.github.io/officer/dev/reference/prop_table.md),
[`table_colwidths()`](https://davidgohel.github.io/officer/dev/reference/table_colwidths.md),
[`table_layout()`](https://davidgohel.github.io/officer/dev/reference/table_layout.md),
[`table_stylenames()`](https://davidgohel.github.io/officer/dev/reference/table_stylenames.md),
[`table_width()`](https://davidgohel.github.io/officer/dev/reference/table_width.md)

## Examples

``` r
table_conditional_formatting(first_row = TRUE, first_column = TRUE)
#> $first_row
#> [1] TRUE
#> 
#> $first_column
#> [1] TRUE
#> 
#> $last_row
#> [1] FALSE
#> 
#> $last_column
#> [1] FALSE
#> 
#> $no_hband
#> [1] FALSE
#> 
#> $no_vband
#> [1] TRUE
#> 
#> attr(,"class")
#> [1] "table_conditional_formatting"
```
