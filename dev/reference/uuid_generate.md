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
#> [1] "6d91466a-608b-45eb-a928-d153dfb6aed1"
#> [2] "485973f0-163d-4db0-8212-f4b4f7747a69"
#> [3] "89f65451-15fb-443f-a467-6ddf005c14fa"
#> [4] "ae94bf04-49c0-4841-9fdc-0345cd0e1121"
#> [5] "5a642d9f-531e-44b6-a55c-3badd2161d37"
```
