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
#> [1] "cc633c60-05fa-4a5b-ba56-68172c0d62ff"
#> [2] "7574f8b7-9710-4828-8f13-ba36e0bf2b54"
#> [3] "9314bc12-2386-4b26-a603-8a1fae3c9884"
#> [4] "d76a2ce9-6374-4a0d-b66a-94bd1374167e"
#> [5] "9503aeef-6241-444e-ba8a-762e3e7c3efd"
```
