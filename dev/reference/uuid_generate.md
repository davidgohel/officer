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
#> [1] "2bcd6282-8485-4868-9dc5-d25e1bd4eccf"
#> [2] "289744c4-c0b5-4eff-9e54-df89aef826de"
#> [3] "c49f4f7a-aa65-4b36-a541-314db4559d52"
#> [4] "28dd0c84-67cd-4bdf-aee9-74f853f1248e"
#> [5] "e017fe2e-7f2b-4470-b928-20299efaf433"
```
