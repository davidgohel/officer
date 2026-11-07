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
#> [1] "8f9cba4a-499f-424b-a424-2e9501fdd845"
#> [2] "5a37f5e7-a9b7-4907-828f-bc1fb7f9cc5c"
#> [3] "f28a1e81-d2b2-45ab-a3c3-b5b13e4fb4fe"
#> [4] "cc0a1236-06fc-4be0-8d0a-4ba5dfa45fc1"
#> [5] "b106931e-664a-47d3-8cfb-26fdfbe8a046"
```
