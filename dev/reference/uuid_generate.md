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
#> [1] "28c7bc00-775b-4d3a-bc1e-de033acefc55"
#> [2] "06d3fc45-8b4f-49e6-bf8f-fb184a42b14b"
#> [3] "eb32dd7e-1584-44cb-8314-60a9c0c2b8ac"
#> [4] "04ef06e7-0ed3-48e2-9444-52042c09834b"
#> [5] "db8bb905-81e3-42bc-95bf-f9b8b13cac92"
```
