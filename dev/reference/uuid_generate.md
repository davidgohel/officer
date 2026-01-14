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
#> [1] "a1c321b4-d558-46ef-94b6-12e8a8c80a6b"
#> [2] "31dfde8e-4a20-4512-a6a3-b8bd5c63f7fe"
#> [3] "1ffa15b4-c539-42f8-baeb-c099f5cad808"
#> [4] "797556be-46ad-4059-9ff5-c7d5cc4123b2"
#> [5] "30191cb7-52ef-4513-8e25-8ea9697d81b3"
```
