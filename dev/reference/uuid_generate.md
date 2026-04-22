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
#> [1] "c6e0b6ef-a7f2-444e-8847-2784b34e3d43"
#> [2] "b07826e0-4a7c-4b86-a3e5-cb7b5f606da4"
#> [3] "a3a54c07-59d1-4bf9-992c-b15daa8f4e69"
#> [4] "5b1e0614-3ec1-41f4-b4cd-7dea4f7cb990"
#> [5] "45eac772-44d3-432c-b19a-a55d233fea55"
```
