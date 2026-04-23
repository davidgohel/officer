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
#> [1] "88e6c903-d10c-4262-9a8b-80817c21fab7"
#> [2] "0f243fe1-b24a-423f-939b-f805a553e551"
#> [3] "715bc450-0201-49f6-bde0-131353cd2227"
#> [4] "0b8d7844-7953-4045-b3e0-a215bda7fa19"
#> [5] "adc09bc0-6c7a-4bcf-ae64-3a459dff433b"
```
