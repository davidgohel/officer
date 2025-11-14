# officer url encoder

encode url so that it can be easily decoded when 'officer' write a file
to the disk.

## Usage

``` r
officer_url_encode(x)
```

## Arguments

- x:

  a character vector of URL

## Examples

``` r
officer_url_encode("https://cran.r-project.org/")
#> [1] "https://cran.r-project.org/"
```
