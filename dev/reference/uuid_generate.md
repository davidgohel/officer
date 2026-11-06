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
#> [1] "a246aa0b-1b14-436d-b6db-821e829be4f4"
#> [2] "96962666-4d4f-4fc7-bcc2-8015035c4b0b"
#> [3] "7c6d1a0a-4897-42df-8ce9-71da794bade1"
#> [4] "9fbd367e-8224-4489-801f-851165838360"
#> [5] "deb99d08-98eb-41e4-9b65-cd7965a83b1f"
```
