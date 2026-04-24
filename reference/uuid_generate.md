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
#> [1] "c195618f-fa01-42e0-9d97-2abf29bd0536"
#> [2] "c3b9e2fd-c949-4cb6-b956-9056ad960c31"
#> [3] "60891e6e-32ea-465e-8537-a9ac04f7953d"
#> [4] "a2047b9c-837e-4571-bbee-8f917e42f746"
#> [5] "b6ec4ba0-5a6b-4b03-9d17-3a86e54a14b9"
```
