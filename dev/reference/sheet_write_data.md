# Write data to a sheet

Write a data.frame into a sheet of an xlsx workbook. Multiple calls can
write to different positions on the same sheet.

## Usage

``` r
sheet_write_data(x, data, sheet, start_row = 1L, start_col = 1L)
```

## Arguments

- x:

  rxlsx object

- data:

  a data.frame

- sheet:

  sheet name (must already exist)

- start_row:

  row index where the header will be written (default 1)

- start_col:

  column index where the first column of data will be written (default
  1)

## Examples

``` r
x <- read_xlsx()
x <- add_sheet(x, label = "mysheet")
x <- sheet_write_data(x, data = head(iris, 5), sheet = "mysheet")
x <- sheet_write_data(x, data = head(mtcars, 3), sheet = "mysheet",
  start_col = 7)
print(x, target = tempfile(fileext = ".xlsx"))
```
