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
#> [1] "d9e7f3ea-50f9-416a-905e-dffe6588d5cb"
#> [2] "7b8aee70-28ce-4c7d-8778-caac311425ea"
#> [3] "dffb0244-e26f-4c57-90db-7beee590c90a"
#> [4] "7e8d54a2-f767-42b9-922c-03f90cd650d1"
#> [5] "e74015c2-68a4-43df-9879-7980f504a684"
```
