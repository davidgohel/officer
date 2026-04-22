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
#> [1] "23ace860-ba38-4721-84b9-9b8ef964d004"
#> [2] "f528ada5-f9f5-43b6-bd3d-1716f3653109"
#> [3] "3a3a154e-0a28-40cb-b789-13cf5fa4479d"
#> [4] "6b355f9f-f060-41d2-b3a2-286b86985f55"
#> [5] "caea6c9c-85dc-4004-a6c8-7ae1e17270f5"
```
