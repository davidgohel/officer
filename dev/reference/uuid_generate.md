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
#> [1] "9da2b25d-2bf2-4b7c-b5a8-aa2d8e825e5a"
#> [2] "698b39a3-4b9f-4fae-a086-b1388ad8210a"
#> [3] "8fd632ac-047c-451a-a0e4-e9b41df073cc"
#> [4] "cbc1ef04-cd53-4fd4-b766-e90d2d1fcc7c"
#> [5] "c6e60494-12a9-4e6a-816f-61888ff4c1b7"
```
