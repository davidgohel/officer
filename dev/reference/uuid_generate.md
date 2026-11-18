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
#> [1] "739361d9-3a22-44e5-930a-f6a87997d88b"
#> [2] "2f655c98-1c03-4829-8aff-088c340e57a1"
#> [3] "161678b6-6a61-4f56-97f5-d3857b0a9ee5"
#> [4] "5a687473-df24-44ff-a407-65c2b864d414"
#> [5] "02978fc0-02d6-40d5-abe5-f86ab5e890f3"
```
