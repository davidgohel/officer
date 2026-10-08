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
#> [1] "d7475160-ce3e-4ade-ad89-20b13db049d5"
#> [2] "eebf8503-bc6c-4b0f-b3bb-49c45ae3d4d8"
#> [3] "d156f509-43f7-4977-ae57-a421104c0e4d"
#> [4] "60c67f18-7f8e-4dbf-9a7b-2fc2b2a15aaa"
#> [5] "c0d90ea2-44e9-4448-ac66-a1a863d1cf83"
```
