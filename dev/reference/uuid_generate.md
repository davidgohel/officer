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
#> [1] "8167bfe8-4fc7-479c-b5d0-3387b9144ac0"
#> [2] "2520f84a-30a0-4207-9cc0-4c818ca6ea11"
#> [3] "fc83f2b9-549a-4dae-bb81-02068cf74ba9"
#> [4] "47c29a92-21f9-443e-b394-15ec84e5ed36"
#> [5] "aa7b5759-d68f-409c-81a2-3335ec131d8f"
```
