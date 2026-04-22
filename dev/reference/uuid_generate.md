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
#> [1] "29e3ad23-f650-4f50-a8d6-b42eb62d7997"
#> [2] "0d4058a7-1af9-4480-a78b-18a73bf8d5cb"
#> [3] "7997a4ee-ba7e-48d4-ae4f-dcb1dc9f20c6"
#> [4] "e2d83701-8d9c-4be3-b443-61f6fc175fa6"
#> [5] "fab8601f-d181-4e4e-84ae-140e6bf852f6"
```
