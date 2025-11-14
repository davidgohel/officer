# Select sheet

Set a particular sheet selected when workbook will be edited.

## Usage

``` r
sheet_select(x, sheet)
```

## Arguments

- x:

  rxlsx object

- sheet:

  sheet name

## Examples

``` r
my_ws <- read_xlsx()
my_pres <- add_sheet(my_ws, label = "new sheet")
my_pres <- sheet_select(my_ws, sheet = "new sheet")
print(my_ws, target = tempfile(fileext = ".xlsx") )
```
