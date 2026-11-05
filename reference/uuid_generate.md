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
#> [1] "04beff20-da4e-4784-b1b3-0947020d9a61"
#> [2] "7a4be0f9-9c56-4d1f-be36-9ff9ce4dc830"
#> [3] "cf7501ed-7b2f-4505-a7c0-054365563093"
#> [4] "a8cd1469-b08d-4f1c-bc63-fb60b1942109"
#> [5] "0c937dc9-c169-46c7-b194-11e0a545de73"
```
