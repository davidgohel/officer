# Algorithm for table layout

When a table is displayed in a document, it can either be displayed
using a fixed width or autofit layout algorithm:

- fixed: uses fixed widths for columns. The width of the table is not
  changed regardless of the contents of the cells.

- autofit: uses the contents of each cell and the table width to
  determine the final column widths.

## Usage

``` r
table_layout(type = "autofit")
```

## Arguments

- type:

  'autofit' or 'fixed' algorithm. Default to 'autofit'.

## See also

Other functions for table definition:
[`prop_table()`](https://davidgohel.github.io/officer/dev/reference/prop_table.md),
[`table_colwidths()`](https://davidgohel.github.io/officer/dev/reference/table_colwidths.md),
[`table_conditional_formatting()`](https://davidgohel.github.io/officer/dev/reference/table_conditional_formatting.md),
[`table_stylenames()`](https://davidgohel.github.io/officer/dev/reference/table_stylenames.md),
[`table_width()`](https://davidgohel.github.io/officer/dev/reference/table_width.md)
