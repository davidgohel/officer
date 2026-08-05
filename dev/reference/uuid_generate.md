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
#> [1] "874eb32a-3534-4753-8d1e-25a4602ac58a"
#> [2] "f28160b9-1349-42a8-aae4-dbd391e955c2"
#> [3] "75fa868e-1e27-4f91-b496-3be47e99a53f"
#> [4] "c6935044-ca2d-4aa2-8d6f-4aad6e1decd1"
#> [5] "8690cedb-05ee-422f-ab6b-dc8d82f27769"
```
