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
#> [1] "ba44773f-03bd-4e7b-a5fa-89c3dd23637a"
#> [2] "fc6f03f1-ff9e-48c6-bd1f-9f1a12a5f5bd"
#> [3] "76df9431-3818-4a97-bda1-ef1ed443c382"
#> [4] "0abb6aa5-a26e-4000-846d-e0c397022127"
#> [5] "ab953b96-7219-4f2c-8913-9b52313f46e8"
```
