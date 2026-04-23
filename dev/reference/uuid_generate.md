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
#> [1] "bae783ed-666b-400e-bb75-847f98d9a897"
#> [2] "5546eae0-1847-468a-beeb-61719f6bed6b"
#> [3] "98beac45-f6ce-47b0-bfa1-b370e3874afd"
#> [4] "b159700c-82b4-4160-ace8-b1e67f3059e2"
#> [5] "74c70031-a99e-45a4-968f-5f589492fa65"
```
