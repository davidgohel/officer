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
#> [1] "f8d6526a-553a-4da7-a3ee-d7167bd23df5"
#> [2] "f35ccdfd-f111-4e93-af7a-28e86911939e"
#> [3] "13fe23bb-f71a-4da6-ae76-e7ea9fcd46da"
#> [4] "00ecbc3d-28cb-418b-82c2-10b04f295613"
#> [5] "d3b266e4-f6b5-4335-9bcc-d4660177bf7b"
```
