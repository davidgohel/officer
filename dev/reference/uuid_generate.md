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
#> [1] "3e1790e5-e156-402a-9548-0fc2e0c9faf8"
#> [2] "7b205bde-338b-419d-800d-f0fb77e50ab4"
#> [3] "936a54e5-fd38-4684-afc0-e4211bd37794"
#> [4] "ac3585b9-9928-451c-a5c0-d0d9281ed3cf"
#> [5] "3329ed50-0246-47dd-a4d5-64184a3d34e7"
```
