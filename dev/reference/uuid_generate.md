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
#> [1] "b78ac4b7-2e79-4a60-8aef-571bb4ac0350"
#> [2] "f26975ab-d18c-4201-9fcb-7e4f26c17683"
#> [3] "708434b8-146c-4677-a9e7-cc60f1ace14e"
#> [4] "40c69d53-70e3-4c90-9f80-35a9977ed88b"
#> [5] "45b862f2-423b-4e8e-9f7e-f3fbd74f095a"
```
