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
#> [1] "ad8f0539-1c8b-4a1d-85ea-b0d48d920d12"
#> [2] "0747188b-a85b-42c9-83c5-aeac0642e64b"
#> [3] "886b1822-1910-4ede-b94e-a82e3bcbcaff"
#> [4] "42941b24-8599-4b32-b843-8b03b554596b"
#> [5] "44b860b1-0831-477e-86df-e4daf6e49466"
```
