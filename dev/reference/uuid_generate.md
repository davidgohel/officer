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
#> [1] "257c12d5-ad17-48be-b3c5-bc7208b80296"
#> [2] "d5d87085-7790-4aba-9f5a-638cf99237ce"
#> [3] "0fa23b69-9131-4bee-80ab-19699676b9ab"
#> [4] "7ccebdfc-c392-4bb7-86eb-0ee73c3afe23"
#> [5] "252fb91f-161f-4f34-b45b-03d8ca5f6f80"
```
