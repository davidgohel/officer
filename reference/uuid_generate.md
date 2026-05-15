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
#> [1] "3ed3a06e-d7d1-4b4d-b134-de16798e7951"
#> [2] "d7b12e5a-3abc-49d3-9d29-783af3c5fd80"
#> [3] "a6ad6b72-79b7-4d1b-8498-51df7b816b7a"
#> [4] "53ed8b6d-c388-4933-9e4d-f4e0822c8a8a"
#> [5] "9046631a-90b8-45a8-8155-0de05f137047"
```
