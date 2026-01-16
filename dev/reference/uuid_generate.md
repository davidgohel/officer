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
#> [1] "13077069-ed7e-4016-9315-f11896a23496"
#> [2] "560b73a0-f1cb-4bf9-9d85-425d5462a223"
#> [3] "9878717c-86eb-4d98-b910-63feef986bfa"
#> [4] "8aa5fc82-b86c-42e3-9193-769631ecd46f"
#> [5] "64a85f9f-adee-4074-ac4a-6ab4c8f2abd4"
```
