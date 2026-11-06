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
#> [1] "a95dfdea-458b-4ea9-8905-9d429225b2d8"
#> [2] "d921af42-2f2d-4301-9ec3-3ebd84d7feb5"
#> [3] "655c45f8-b268-4202-81e7-9a53d2e98250"
#> [4] "1433201d-4aea-421e-b981-8f14ef5e06f7"
#> [5] "9b4633b5-a51c-43ec-b16c-088db2f9b94c"
```
