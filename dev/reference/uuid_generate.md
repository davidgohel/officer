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
#> [1] "dab4df54-63ae-435b-839a-92c76da454ce"
#> [2] "56a730ad-aa82-4b44-932f-9fa9e18632e3"
#> [3] "c4bfa818-d4eb-46ad-9d4e-c4998febbe2d"
#> [4] "7d6ed42d-50b6-4a12-9e66-482921d53ed9"
#> [5] "1e4a098c-a2ce-4ad6-a788-1f0c2dda64e5"
```
