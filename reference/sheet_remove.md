# Remove a sheet

Remove a sheet from an xlsx workbook, deleting the worksheet XML, its
relationship file and the content-type override.

## Usage

``` r
sheet_remove(x, sheet)
```

## Arguments

- x:

  rxlsx object

- sheet:

  name of the sheet to remove

## Value

the rxlsx object (invisibly)

## Examples

``` r
wb <- read_xlsx()
wb <- add_sheet(wb, "kept")
default_name <- sheet_names(wb)[1]
wb <- sheet_remove(wb, sheet = default_name)
print(wb, target = tempfile(fileext = ".xlsx"))
```
