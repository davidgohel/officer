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
#> [1] "26c60b4a-8c8f-49f4-9cdd-57cff62c4b08"
#> [2] "4cceea54-a8e6-4d80-8930-dec807ff0df2"
#> [3] "81597456-57c7-4b0f-bb6d-e8f4ca584d76"
#> [4] "dffeba08-e79a-495f-9a0d-99065eab83af"
#> [5] "73c9ccc7-8c7a-4de2-80b7-33a40386ce21"
```
