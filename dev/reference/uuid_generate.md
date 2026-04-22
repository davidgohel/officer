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
#> [1] "1c7487ef-4e08-461e-8bd7-f3974acaa973"
#> [2] "53d36e50-7a9c-4134-bb14-b099a4145a0a"
#> [3] "97a507d0-223a-47b3-bc36-8059d1425bba"
#> [4] "62f1e10e-ed2a-4c5c-953e-a2ccc52624b2"
#> [5] "1e16f138-95d6-46cd-9e46-77f7be172d19"
```
