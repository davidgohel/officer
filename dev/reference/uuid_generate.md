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
#> [1] "7eaac683-978a-4c8c-9ac3-730c355f2a31"
#> [2] "f77d474f-5aed-4f06-80c8-12e43e882c9a"
#> [3] "6d159177-574c-4725-808d-ac2b5ee09e1d"
#> [4] "4a2a8b8a-226a-4a32-bc5e-2d0bf3b6bbf9"
#> [5] "e869cced-748a-4efe-99e3-c74b6594b542"
```
