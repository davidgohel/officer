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
#> [1] "69aee10f-86d9-4994-96f3-a6441b9e6836"
#> [2] "373c0df8-c7e9-4550-a8c6-2730b885a6da"
#> [3] "16f8edbe-273d-4e91-a48b-f75c8d59327b"
#> [4] "3c216ee7-904a-485f-8de4-b835a3fcfd7b"
#> [5] "8627cf76-818f-4716-9653-a9b16f9c4627"
```
