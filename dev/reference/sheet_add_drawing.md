# Add a drawing to an Excel sheet

Add a graphical element into a sheet of an xlsx workbook. This is a
generic function dispatching on `value`. Methods are provided by
extension packages 'mschart' and 'rvg'.

Use
[`sheet_write_data()`](https://davidgohel.github.io/officer/dev/reference/sheet_write_data.md)
to write data into the sheet before or after adding a drawing.

## Usage

``` r
sheet_add_drawing(x, value, sheet, ...)
```

## Arguments

- x:

  rxlsx object created by
  [`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md)

- value:

  object to add (dispatched to the appropriate method)

- sheet:

  sheet name (must already exist, see
  [`add_sheet()`](https://davidgohel.github.io/officer/dev/reference/add_sheet.md))

- ...:

  additional arguments passed to methods

## Value

the rxlsx object (invisibly)

## See also

[`read_xlsx()`](https://davidgohel.github.io/officer/dev/reference/read_xlsx.md),
[`add_sheet()`](https://davidgohel.github.io/officer/dev/reference/add_sheet.md),
[`sheet_write_data()`](https://davidgohel.github.io/officer/dev/reference/sheet_write_data.md)
