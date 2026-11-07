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
#> [1] "eff72ba1-4a91-4cde-89b3-e92b8cda36f8"
#> [2] "7ea891a0-e87e-46d2-9316-c0b02c62a4b3"
#> [3] "0c042ee7-bbb9-4752-bd01-171ec5f44fa4"
#> [4] "eb5d055d-dfea-452a-a0ce-558c6995b125"
#> [5] "d28bc155-a29a-4382-abe1-d08802ca0864"
```
