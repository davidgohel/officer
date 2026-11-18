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
#> [1] "ea42df04-cb4b-4ed5-8663-5b28a851f37b"
#> [2] "157908dd-6803-4d7c-b31b-668dd15db011"
#> [3] "bed14649-8811-4779-b1eb-c71176db4f0e"
#> [4] "6ae9cf37-317a-4b3d-b65f-bf202e27be6c"
#> [5] "8bab49bc-a426-4da8-be0d-f43a26319686"
```
