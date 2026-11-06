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
#> [1] "861e4b54-99e4-4c9d-86f5-76154ef460f3"
#> [2] "dcc512ad-1267-4dcb-9f87-5b6986905df9"
#> [3] "da2a00a6-072d-4c91-b156-fc5d6154f572"
#> [4] "d88d25aa-b33e-4ecf-a320-e63e57e8f6d6"
#> [5] "96b72377-4d28-4012-ba23-ad5647e8a1eb"
```
