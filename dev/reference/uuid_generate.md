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
#> [1] "a871c0bc-fe69-4b0d-a7c5-a620e52cef42"
#> [2] "a717e4f8-3a92-4ff1-bd27-986169cac532"
#> [3] "a1474f75-e0e7-4db7-a777-cdc0b2ac4cde"
#> [4] "a60e86c6-f86c-4500-8e29-7ac53848373f"
#> [5] "383bdeb3-3f13-46ef-bcc3-211d9cf22ec7"
```
