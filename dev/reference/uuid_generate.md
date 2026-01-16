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
#> [1] "3eb13807-567f-475b-a0b9-8fce24365da3"
#> [2] "90eae051-226a-4fb3-b35e-0196523fbc2d"
#> [3] "265bc66f-cbf3-4388-a28c-b35d56264f8e"
#> [4] "84601846-3802-42fa-a3a9-69db7051e0ba"
#> [5] "967cf6cf-b4a0-4de1-b1ca-6387d3b43f37"
```
