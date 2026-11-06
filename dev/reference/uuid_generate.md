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
#> [1] "f0eb5aa1-e3c7-4fe5-821b-20a15333e517"
#> [2] "5689c0ec-beb1-4e8b-9623-54837d29f606"
#> [3] "eb8479bc-8a45-4cc1-97d3-a0871e2cbaef"
#> [4] "0e5ff628-9d97-4e2b-807e-c33b4a352e96"
#> [5] "89e38d0e-e25c-41a1-bd07-affff987e52d"
```
