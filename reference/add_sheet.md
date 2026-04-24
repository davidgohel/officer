# Add a sheet

Add a sheet into an xlsx worksheet.

## Usage

``` r
add_sheet(x, label)
```

## Arguments

- x:

  rxlsx object

- label:

  sheet label

## Details

[`read_xlsx()`](https://davidgohel.github.io/officer/reference/read_xlsx.md)
returns a workbook that already contains one default sheet shipped with
the template (named `"Sheet1"` or `"Feuil1"` depending on the locale).
`add_sheet()` is purely additive: the default sheet is kept as-is.
Remove it explicitly with
[`sheet_remove()`](https://davidgohel.github.io/officer/reference/sheet_remove.md)
if it is not wanted.

## Examples

``` r
my_ws <- read_xlsx()
my_pres <- add_sheet(my_ws, label = "new sheet")
```
