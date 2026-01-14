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
#> [1] "c146b560-1d7a-4f69-ac6f-bfec38933473"
#> [2] "e68b543a-3fde-4a5a-95b3-2bb730cce8e3"
#> [3] "69746208-3f1b-4c43-96ab-a917a855d0ad"
#> [4] "2bd3bbb8-fd51-4f2c-a3c1-1e3ab9d27e5c"
#> [5] "28c0bffd-5526-40c8-908f-a380b0055b9a"
```
