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
default_name <- wb$worksheets$sheet_names()[1]
# touch the default sheet first so add_sheet does not auto-drop it
wb <- sheet_write_data(wb, head(iris, 2), sheet = default_name)
wb <- add_sheet(wb, "kept")
wb <- sheet_remove(wb, sheet = default_name)
print(wb, target = tempfile(fileext = ".xlsx"))
```
