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
#> [1] "458cb8a4-d56e-4913-b1f9-1aed1f9d5c1a"
#> [2] "1ea68851-71b5-47eb-815e-e38f55571fb1"
#> [3] "7453b822-3f51-4adc-a495-097df1eddd24"
#> [4] "c414a69d-4848-4305-92f3-bfbf0ebbc28d"
#> [5] "e5811521-2cfb-4eb3-b764-2a777db13e56"
```
