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
#> [1] "3cf10985-ce01-4d69-8fa9-18b890587fd1"
#> [2] "b5f0b2c3-c297-4ee2-8def-cee2c1e0b08f"
#> [3] "695a4f34-0c58-4283-a5c4-8c84075caff0"
#> [4] "42640993-80e1-4c7b-8692-c3e0336835a9"
#> [5] "cbebe521-d179-4260-b5b3-7c8d7b0c6db9"
```
