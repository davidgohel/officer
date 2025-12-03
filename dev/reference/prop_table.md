# Table properties

Define table properties such as fixed or autofit layout, table width in
the document, eventually column widths.

## Usage

``` r
prop_table(
  style = NA_character_,
  layout = table_layout(),
  width = table_width(),
  stylenames = table_stylenames(),
  colwidths = table_colwidths(),
  tcf = table_conditional_formatting(),
  align = "center",
  word_title = NULL,
  word_description = NULL
)
```

## Arguments

- style:

  table style to be used to format table

- layout:

  layout defined by
  [`table_layout()`](https://davidgohel.github.io/officer/dev/reference/table_layout.md),

- width:

  table width in the document defined by
  [`table_width()`](https://davidgohel.github.io/officer/dev/reference/table_width.md)

- stylenames:

  columns styles defined by
  [`table_stylenames()`](https://davidgohel.github.io/officer/dev/reference/table_stylenames.md)

- colwidths:

  column widths defined by
  [`table_colwidths()`](https://davidgohel.github.io/officer/dev/reference/table_colwidths.md)

- tcf:

  conditional formatting settings defined by
  [`table_conditional_formatting()`](https://davidgohel.github.io/officer/dev/reference/table_conditional_formatting.md)

- align:

  table alignment (one of left, center or right)

- word_title:

  alternative text for Word table (used as title of the table)

- word_description:

  alternative text for Word table (used as description of the table)

## See also

Other functions for table definition:
[`table_colwidths()`](https://davidgohel.github.io/officer/dev/reference/table_colwidths.md),
[`table_conditional_formatting()`](https://davidgohel.github.io/officer/dev/reference/table_conditional_formatting.md),
[`table_layout()`](https://davidgohel.github.io/officer/dev/reference/table_layout.md),
[`table_stylenames()`](https://davidgohel.github.io/officer/dev/reference/table_stylenames.md),
[`table_width()`](https://davidgohel.github.io/officer/dev/reference/table_width.md)

## Examples

``` r
prop_table()
#> $style
#> [1] NA
#> 
#> $layout
#> $type
#> [1] "autofit"
#> 
#> attr(,"class")
#> [1] "table_layout"
#> 
#> $width
#> $width
#> [1] 1
#> 
#> $unit
#> [1] "pct"
#> 
#> attr(,"class")
#> [1] "table_width"
#> 
#> $colsizes
#> $widths
#> NULL
#> 
#> attr(,"class")
#> [1] "table_colwidths"
#> 
#> $stylenames
#> $stylenames
#> list()
#> 
#> attr(,"class")
#> [1] "table_stylenames"
#> 
#> $tcf
#> $first_row
#> [1] TRUE
#> 
#> $first_column
#> [1] FALSE
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
#> 
#> $align
#> [1] "center"
#> 
#> $word_title
#> NULL
#> 
#> $word_description
#> NULL
#> 
#> attr(,"class")
#> [1] "prop_table"
to_wml(prop_table())
#> [1] "<w:tblPr><w:tblLayout w:type=\"autofit\"/><w:jc w:val=\"center\"/><w:tblW w:type=\"pct\" w:w=\"5000\"/><w:tblLook w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"0\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr>"
```
