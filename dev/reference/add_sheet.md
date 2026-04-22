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

[`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)
returns a workbook that already contains one default sheet shipped with
the template (named `"Sheet1"` or `"Feuil1"` depending on the locale).
When the user adds their first sheet with `add_sheet()`, that default
sheet is silently dropped if it still looks empty (no cell content, no
drawing attached), so the resulting xlsx does not start with a stray
empty tab. A default sheet that has been written into (via
[`sheet_write_data()`](https://davidgohel.github.io/officer/dev/reference/sheet_write_data.md)
or
[`sheet_add_drawing()`](https://davidgohel.github.io/officer/dev/reference/sheet_add_drawing.md))
is kept.

To bypass the auto-drop, touch the default sheet before adding any other
sheet:

    wb <- read_xlsx()
    wb <- sheet_write_data(wb, head(iris),
                           sheet = wb$worksheets$sheet_names()[1])
    wb <- add_sheet(wb, "second")   # both sheets kept

To remove a sheet explicitly later, use
[`sheet_remove()`](https://davidgohel.github.io/officer/dev/reference/sheet_remove.md).

## Examples

``` r
my_ws <- read_xlsx()
my_pres <- add_sheet(my_ws, label = "new sheet")
```
