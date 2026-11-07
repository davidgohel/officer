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
#> [1] "38328034-96c5-47a9-a38a-ad431663ad4d"
#> [2] "e105cd9f-d9f7-4f23-b089-1ed007d8110d"
#> [3] "8c1c7068-9d54-469f-ae5c-27ab4618d23e"
#> [4] "43b4e033-c306-4a8a-9ae0-de7a6a421605"
#> [5] "d032d425-2ede-47ea-ba0c-75abe5a6789b"
```
