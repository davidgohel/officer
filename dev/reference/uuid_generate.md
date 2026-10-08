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
#> [1] "746d84ef-d106-4076-b1b1-ed03d9f23f57"
#> [2] "6d01cd66-6743-4c64-8a3b-4d0a1e907629"
#> [3] "af4acb8b-0b4a-4ed3-94bb-b718ffe7605f"
#> [4] "d48b0eac-bf49-4237-b53e-d85d5949adb2"
#> [5] "a6e1c169-6ea0-4aa9-9244-949b46f3b614"
```
