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
#> [1] "fb7c53d5-18e7-4561-bd0a-4d4c737d38f0"
#> [2] "089d1fa9-dec1-4e01-8720-b633fddf5b9e"
#> [3] "4244a11c-3efe-41a2-9b31-f967fc1108cd"
#> [4] "15385bb1-a12e-4696-b9fc-7c4c52f1e593"
#> [5] "0acab81c-b002-48de-a1fc-dbe4d681efee"
```
