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
#> [1] "72243c4f-652d-48c6-b720-9291af844eef"
#> [2] "ac03eef1-be44-4eac-b386-cd8e264205fa"
#> [3] "9e37f098-6fc2-473d-9f84-70748e6fc869"
#> [4] "7d265a43-1d8b-45be-b575-fe3bb93c094b"
#> [5] "f5eaed8a-a86c-404f-989a-161b7301972a"
```
