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
#> [1] "e7990457-8853-4d73-8689-b3e335462a49"
#> [2] "a41b102c-69c0-4291-b4ca-67ccb87b1f12"
#> [3] "3fe1501a-cd0b-4640-aeab-3d262aa951f6"
#> [4] "3d23d9a5-8357-4be4-b04c-d3618e39afb9"
#> [5] "6781ad5d-9b35-40ce-b93f-e9ff9b8c9d41"
```
