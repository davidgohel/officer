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
#> [1] "0ac7c190-c850-4d74-b093-265b5743d23b"
#> [2] "0c9c5a66-b1d5-4618-8c43-dcf53e212cd5"
#> [3] "cd46fbf5-0971-4b61-9aec-604c52ca48bf"
#> [4] "e9ed80ab-40fa-4da6-bd09-3e049980ffb2"
#> [5] "71e5a11a-870f-43fe-b344-9c641b28b366"
```
