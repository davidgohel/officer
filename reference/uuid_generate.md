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
#> [1] "51f32434-b848-466f-ab61-b37d787b90b3"
#> [2] "45873e30-6c7c-49b3-aa0f-818575128e46"
#> [3] "4a314373-d65a-42ac-88fb-aba3e4708af9"
#> [4] "d91ea4c5-265d-4ebe-bc39-dce07bb4f647"
#> [5] "0de2742a-2da6-4c4c-9f87-435fb4d4cdaf"
```
