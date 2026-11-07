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
#> [1] "52335838-97f2-447e-9e2a-96f18df5f890"
#> [2] "6fae563e-abeb-4a13-b557-2b0d3dd0a642"
#> [3] "88216c09-3583-4e66-8aa8-c6d04a405e5d"
#> [4] "2831a082-cf07-4d29-a0bf-6d303ccc844c"
#> [5] "477ae308-9a13-4da3-9532-4c9afae167cc"
```
