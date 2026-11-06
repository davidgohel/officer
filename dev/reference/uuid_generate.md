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
#> [1] "8205be08-9b84-43bc-832a-6a2cf5f55027"
#> [2] "92b12104-cdd4-4f5d-8b53-fc343dc859d1"
#> [3] "82b55527-7f77-427a-9807-ecdbfdc2021e"
#> [4] "1ffe1900-2544-4363-ac8b-ca3fb948e2aa"
#> [5] "66c14623-bf8e-40bc-a5aa-d1a03afa4dbd"
```
