# Create an 'Excel' document object

Read and import an xlsx file as an R object representing the document.
This function is experimental.

## Usage

``` r
read_xlsx(path = NULL)

# S3 method for class 'rxlsx'
length(x)

# S3 method for class 'rxlsx'
print(x, target = NULL, ...)
```

## Arguments

- path:

  path to the xlsx file to use as base document.

- x:

  an rxlsx object

- target:

  path to the xlsx file to write

- ...:

  unused

## Examples

``` r
read_xlsx()
#> xlsx document with 1 sheet(s):
#> [1] "Feuil1"
x <- read_xlsx()
print(x, target = tempfile(fileext = ".xlsx"))
```
