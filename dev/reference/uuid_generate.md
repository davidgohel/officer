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
#> [1] "7ff10bee-43e2-4892-adc8-ac1f0dad35d5"
#> [2] "a8217891-588c-495e-b362-d6e6810680dd"
#> [3] "d2436be1-8d95-4978-b5cd-52d7322b93dd"
#> [4] "3710d6d0-6bf8-415f-badf-9f03516b559c"
#> [5] "2de1afec-e612-4b14-8429-21359220c320"
```
