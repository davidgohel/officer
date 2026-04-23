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
#> [1] "a92f8fed-e1c7-4626-8de1-4751c011870b"
#> [2] "ed9fd78a-1b62-4f77-ace2-dbe23c86d775"
#> [3] "f8923df8-1d38-4588-9c90-ffb28cb46624"
#> [4] "c1e265df-a4b6-4588-bef2-38847e7faf4f"
#> [5] "921f7941-e17c-40a3-b784-9d2656199f24"
```
