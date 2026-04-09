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
#> [1] "82ebafdf-d5c1-486a-8c14-9a37ea577106"
#> [2] "7c589161-6072-46c9-94da-e8f7f0898c2a"
#> [3] "da1fb289-471d-4b58-aecf-57631c0702e8"
#> [4] "e987fea3-23c4-4938-82c5-e9d312686e09"
#> [5] "87681a4c-def7-447e-9f4c-7c3a89c04113"
```
