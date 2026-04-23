# Sheet names of an xlsx workbook

Return the sheet names of an `rxlsx` object, in the order they appear in
the workbook.

## Usage

``` r
sheet_names(x)
```

## Arguments

- x:

  an rxlsx object (created by
  [`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)).

## Value

A character vector of sheet names.

## Examples

``` r
wb <- read_xlsx()
sheet_names(wb)
#> [1] "Feuil1"

wb <- add_sheet(wb, label = "new sheet")
sheet_names(wb)
#> [1] "Feuil1"    "new sheet"
```
