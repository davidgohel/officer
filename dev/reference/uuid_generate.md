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
#> [1] "a7e3848f-407e-4cd8-8e8e-777a0fa4a86a"
#> [2] "400f667c-6b5f-4674-b243-d18970b77e11"
#> [3] "2d4bdfcf-32e3-48a9-a3c5-2bb301552180"
#> [4] "5a9a95ed-dacb-4c03-bfd7-e98bb203b38f"
#> [5] "ecb08400-fbc0-458f-b5af-85cc9f256e47"
```
