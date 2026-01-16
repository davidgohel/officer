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
#> [1] "66513833-afc8-4d76-9f3b-da9f73819402"
#> [2] "a279ca71-1789-4d75-9218-4dd285dfdfbc"
#> [3] "cd4ce02b-f942-4267-8828-4ae9adaa7582"
#> [4] "26405bbe-3b73-44b3-a2a8-ce3383be42a2"
#> [5] "110afd15-c390-44e0-9daf-9431d9d350a4"
```
