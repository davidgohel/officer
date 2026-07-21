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
#> [1] "59b81b66-db62-4d70-8b23-7fb806784c1b"
#> [2] "edfce507-ff82-4df4-b69b-47ac169f9015"
#> [3] "4a1bf6b1-2f52-413b-a45f-c0aba60f9b8a"
#> [4] "67c687f1-9cc0-489c-b657-c81e1ecd3d4e"
#> [5] "d08eb37f-c129-436c-aa8e-eafa6951f6d0"
```
