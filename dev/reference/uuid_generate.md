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
#> [1] "618e5b72-0885-4914-bb61-72dc916e4e81"
#> [2] "819d53af-1dae-4bb5-a15d-725768638c24"
#> [3] "39e0f5aa-7b55-40fc-aa04-1f032526d6be"
#> [4] "09609084-6172-4e99-a64a-4157d1467287"
#> [5] "783bdf77-0e6e-4429-87e1-110ffc8f4077"
```
