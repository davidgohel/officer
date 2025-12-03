# Preferred width for a table

Define the preferred width for a table.

## Usage

``` r
table_width(width = 1, unit = "pct")
```

## Arguments

- width:

  value of the preferred width of the table.

- unit:

  unit of the width. Possible values are 'in' (inches) and 'pct'
  (percent)

## Word

All widths in a table are considered preferred because widths of columns
can conflict and the table layout rules can require a preference to be
overridden.

## See also

Other functions for table definition:
[`prop_table()`](https://davidgohel.github.io/officer/dev/reference/prop_table.md),
[`table_colwidths()`](https://davidgohel.github.io/officer/dev/reference/table_colwidths.md),
[`table_conditional_formatting()`](https://davidgohel.github.io/officer/dev/reference/table_conditional_formatting.md),
[`table_layout()`](https://davidgohel.github.io/officer/dev/reference/table_layout.md),
[`table_stylenames()`](https://davidgohel.github.io/officer/dev/reference/table_stylenames.md)
