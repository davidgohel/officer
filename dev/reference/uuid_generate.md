# generates unique identifiers

generates unique identifiers based on
[`uuid::UUIDgenerate()`](https://rdrr.io/pkg/uuid/man/UUIDgenerate.html).

## Usage

``` r
uuid_generate(n = 1, ...)
```

## Arguments

- n:

  integer, number of unique identifiers to generate.

- ...:

  arguments sent to
  [`uuid::UUIDgenerate()`](https://rdrr.io/pkg/uuid/man/UUIDgenerate.html)

## Examples

``` r
uuid_generate(n = 5)
#> [1] "85dd2d3d-dc6f-40d7-a3b9-872dbc9ada54"
#> [2] "76b74537-154d-4a46-a39d-e9f8dacfdf29"
#> [3] "769ec201-b3d0-46f1-b5cd-eee743aa70ad"
#> [4] "26cb35e7-361a-44d4-954e-ce600b6b0bd7"
#> [5] "46bda0b6-b512-463e-bccf-498336656957"
```
