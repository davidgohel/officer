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
#> [1] "a0926447-e2b5-43d4-8705-c74d2f0c1792"
#> [2] "a54ccb03-6bf6-44eb-846b-bc414ecf2a8c"
#> [3] "c83c84a7-4c7c-4b0b-a8c4-b9f8552cda13"
#> [4] "f8a4731e-db62-4157-9bfd-e8389abcbee8"
#> [5] "340a8cf9-2132-425d-adb8-f39613159de2"
```
