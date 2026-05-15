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
#> [1] "884e94bb-1d23-4cc2-8bd6-0e3be22a3220"
#> [2] "25de84df-3a29-49fd-9abf-e71225be56c8"
#> [3] "afa0344b-21df-4a10-819c-68a247457c7f"
#> [4] "1cd4ee3e-7494-4c28-9e73-658cbda8e900"
#> [5] "bab9b49d-9e33-46b8-92ea-32e55f3ce7ad"
```
