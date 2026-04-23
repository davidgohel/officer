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
#> [1] "09d6b91c-d81e-46e0-92c1-eaf1832a1cf4"
#> [2] "fa691ef8-71e6-4ac0-b0f8-11028f1d15d4"
#> [3] "c5432163-3ef6-49a1-9a7a-3df71fc40255"
#> [4] "e5b561b8-7f32-4332-95f3-d5daa40b7b89"
#> [5] "69beef8d-2298-4aad-b844-eb9c31c5f613"
```
