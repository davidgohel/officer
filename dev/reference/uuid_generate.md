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
#> [1] "d7c81e93-427d-46e3-bf19-1e2474b6eecc"
#> [2] "e245d473-6fe0-45dd-be4e-f296ce918081"
#> [3] "75ddfb4e-dbc0-436b-86e7-b9700e77d35f"
#> [4] "29d8bd30-082f-4f9f-a854-2ef92f15a189"
#> [5] "a5dcdedd-a1f1-40fd-9ee7-8f5f963cb19d"
```
