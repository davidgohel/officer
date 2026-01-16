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
#> [1] "11ca3ff5-e25e-4a98-9198-822cf2ba6701"
#> [2] "3a784ba2-5a42-4f67-9e0c-5e94bf5aa8de"
#> [3] "cb872fc6-9dcd-4ca0-9a46-4143ba58a7ab"
#> [4] "f1bce99d-f006-4b49-a5a3-93251dd813ed"
#> [5] "82339ce0-3f7a-46a4-b12a-f2e93b830e41"
```
