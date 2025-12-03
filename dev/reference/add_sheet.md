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

## Examples

``` r
my_ws <- read_xlsx()
my_pres <- add_sheet(my_ws, label = "new sheet")
```
