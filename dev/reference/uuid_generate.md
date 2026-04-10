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
#> [1] "9ee67244-c2ee-44d1-927d-6ee2c7b9f4de"
#> [2] "8b8a5e17-28a7-4481-af19-b3305f6a75ae"
#> [3] "b39cb62d-f48f-42a8-a066-3bc4ad64d803"
#> [4] "fc265e7b-727f-4f3b-bf36-f9b1807ce7b8"
#> [5] "e13c8f5c-af73-4c8c-bccc-b1545f7e16eb"
```
