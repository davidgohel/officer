# Table block

Create a representation of a table

## Usage

``` r
block_table(x, header = TRUE, properties = prop_table(), alignment = NULL)
```

## Arguments

- x:

  a data.frame to add as a table

- header:

  display header if TRUE

- properties:

  table properties, see
  [`prop_table()`](https://davidgohel.github.io/officer/reference/prop_table.md).
  Table properties are not handled identically between Word and
  PowerPoint output format. They are fully supported with Word but for
  PowerPoint (which does not handle as many things as Word for tables),
  only conditional formatting properties are supported.

- alignment:

  alignment for each columns, 'l' for left, 'r' for right and 'c' for
  center. Default to NULL.

## See also

[`prop_table()`](https://davidgohel.github.io/officer/reference/prop_table.md)

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/reference/block_section.md),
[`block_toc()`](https://davidgohel.github.io/officer/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/reference/unordered_list.md)

## Examples

``` r
block_table(x = head(iris))
#> 'data.frame':    6 obs. of  5 variables:
#>  $ Sepal.Length: num  5.1 4.9 4.7 4.6 5 5.4
#>  $ Sepal.Width : num  3.5 3 3.2 3.1 3.6 3.9
#>  $ Petal.Length: num  1.4 1.4 1.3 1.5 1.4 1.7
#>  $ Petal.Width : num  0.2 0.2 0.2 0.2 0.2 0.4
#>  $ Species     : Factor w/ 3 levels "setosa","versicolor",..: 1 1 1 1 1 1

block_table(x = mtcars, header = TRUE,
  properties = prop_table(
    tcf = table_conditional_formatting(
      first_row = TRUE, first_column = TRUE)
  ))
#> 'data.frame':    32 obs. of  11 variables:
#>  $ mpg : num  21 21 22.8 21.4 18.7 18.1 14.3 24.4 22.8 19.2 ...
#>  $ cyl : num  6 6 4 6 8 6 8 4 4 6 ...
#>  $ disp: num  160 160 108 258 360 ...
#>  $ hp  : num  110 110 93 110 175 105 245 62 95 123 ...
#>  $ drat: num  3.9 3.9 3.85 3.08 3.15 2.76 3.21 3.69 3.92 3.92 ...
#>  $ wt  : num  2.62 2.88 2.32 3.21 3.44 ...
#>  $ qsec: num  16.5 17 18.6 19.4 17 ...
#>  $ vs  : num  0 0 1 1 0 1 0 1 1 1 ...
#>  $ am  : num  1 1 1 0 0 0 0 0 0 0 ...
#>  $ gear: num  4 4 4 3 3 3 3 4 4 4 ...
#>  $ carb: num  4 4 1 1 2 1 4 2 2 4 ...
```
