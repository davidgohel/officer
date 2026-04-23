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
#> [1] "5a82c497-e3b6-4328-88c3-87c42f5ac0db"
#> [2] "a0a3db12-0307-4f92-a88d-cb77c19c6272"
#> [3] "f7480aa5-d41b-4700-9fdf-50115ea20598"
#> [4] "fa610d00-5a79-45ad-a192-31b15965de99"
#> [5] "64b94bc8-262e-4309-80f9-316d8401c370"
```
