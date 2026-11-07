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
#> [1] "15ddcdde-911e-4a2f-a56d-d35e3269c839"
#> [2] "e4d17043-4f79-49c9-8f61-f228168f3fbb"
#> [3] "f1e2147b-12c2-44c5-8e53-6e413c71874b"
#> [4] "91e8cc29-bc4f-4a96-b283-5435876ee416"
#> [5] "3e30d173-9284-4dfd-9fe3-e3647f52b0b1"
```
