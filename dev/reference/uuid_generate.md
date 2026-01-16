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
#> [1] "973efd52-53e6-4f4c-b577-0a8fef16846b"
#> [2] "c9d37196-a215-476c-bee0-2cb3aaebf6ce"
#> [3] "c85524bc-c6f4-4e69-9b69-0e334a90aded"
#> [4] "3f49587a-afb4-4f93-8d6b-46358471b608"
#> [5] "5f9ddcc4-dd1a-4c61-be13-9bcad4371308"
```
