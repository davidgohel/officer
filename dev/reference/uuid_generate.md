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
#> [1] "1456649e-71d8-4a5b-b7e3-fd90881d9317"
#> [2] "31a0db91-18b3-4d2e-9e96-0ff4c43c0910"
#> [3] "0e8eb5e0-6c36-4de6-a152-17d77e3edf4c"
#> [4] "16814f77-64e4-4998-9139-e3e16c485768"
#> [5] "ac66e75f-c7be-4ebd-bd71-8754cc356968"
```
