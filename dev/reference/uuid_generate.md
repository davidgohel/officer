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
#> [1] "a3209f6e-b361-4aff-91f8-474686edeb31"
#> [2] "aa689124-21f3-4de4-8f91-cc642b83b8eb"
#> [3] "9fcf7e95-766f-4bd5-a560-8872d56e8921"
#> [4] "cf671bca-de39-470a-8916-e70561415902"
#> [5] "6e890bf2-1a25-46a4-9c8e-732c2ec6ffc5"
```
