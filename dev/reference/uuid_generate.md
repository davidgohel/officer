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
#> [1] "37cd9aff-de63-4f11-91fe-a5e2f9c872fc"
#> [2] "022e09f3-c17a-402f-b278-5de41d7c9c37"
#> [3] "0ed6b8d4-2b75-45e8-a2c3-65d5c1ffcb9e"
#> [4] "cd72e4c1-4c71-4116-a369-bc4b8d705980"
#> [5] "1675ec90-e0f9-4ff9-91d3-266e2d8384c9"
```
