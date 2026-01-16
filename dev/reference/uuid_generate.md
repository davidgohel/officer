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
#> [1] "beda2665-c659-409a-a67f-fb9d46fdac83"
#> [2] "142cf039-523a-43da-8402-7ee292b9f091"
#> [3] "45b0ee83-2db9-4155-8bb1-b771b63624a7"
#> [4] "9f405975-ed84-48e9-b91f-672a204801ee"
#> [5] "a9442be4-a913-4107-947e-0f0aa41f71b8"
```
