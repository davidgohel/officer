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
#> [1] "88ad9aea-250f-40c4-b9fa-c3532b7cd752"
#> [2] "1d37e5b8-7dd9-42a9-8751-4b6d8b8c109e"
#> [3] "c669767e-35cd-4f9d-a4e5-0abe4496f770"
#> [4] "3c40f07a-265d-4f33-bfea-0b95dcea8668"
#> [5] "befd32cc-6530-4aeb-a9f1-4af4122cb420"
```
