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
#> [1] "3bf0d9f2-ac51-49d7-b86c-5ae24cdf38bd"
#> [2] "ee4fb1ff-62fa-47ea-9485-5c8eca760a06"
#> [3] "350536c4-289c-487c-a9ce-d9cb8e474dfb"
#> [4] "73ac7082-a920-46c0-954c-38de432efa0c"
#> [5] "0bc8187d-1da4-4cca-bf6a-be16b2f2530e"
```
