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
#> [1] "627ccdf2-d8e7-4c4c-a680-fac7e93e01d6"
#> [2] "18d0735d-b850-4cc4-becc-a489971df920"
#> [3] "457bd23d-8ce8-45af-a0cf-a813409a5e6c"
#> [4] "14ef62dc-050f-4367-9fc6-2ecf1d493a59"
#> [5] "8e67dd45-18db-4ecb-8bc3-905c281be659"
```
