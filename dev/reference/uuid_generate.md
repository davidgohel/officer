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
#> [1] "445ce676-e1b0-4c93-9075-fa368469450d"
#> [2] "58fe2502-8476-47dc-b2ee-7dadcd18b982"
#> [3] "1abb59f9-e49e-4f47-9ee9-b3f232277e40"
#> [4] "2271704e-0617-4dea-a7ee-ea3ef78378fa"
#> [5] "a033c1bf-493a-445d-b990-289fc1f66bc6"
```
