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
#> [1] "62b7d8b6-04b4-44d9-97e5-61e28301a211"
#> [2] "1064b195-c418-4cb1-be0c-8678a520a46a"
#> [3] "adb9a4ac-abc4-4f7e-9a46-feabadadc17c"
#> [4] "76fd170b-d51d-476c-9ede-53e9a71f966d"
#> [5] "aed92730-33d7-4cfa-9d31-858f3ebc53c0"
```
