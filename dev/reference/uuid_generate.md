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
#> [1] "9f5abb5b-9711-41ac-b5ab-6d79892d18a5"
#> [2] "af412ea4-765e-47bc-a8f3-4a3e7c330b98"
#> [3] "270a258f-2cf9-4b54-9343-cff5e20e7fef"
#> [4] "0446d030-f733-4977-bbf7-588998094cab"
#> [5] "cd6ee9f1-13c1-4447-98fd-2529010575b1"
```
