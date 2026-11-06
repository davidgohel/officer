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
#> [1] "4f1bb8cb-b50e-4570-8474-ffe41685acf5"
#> [2] "5549804c-3925-424a-a62a-15f80cf85f1b"
#> [3] "c0815b75-deeb-41e4-b576-56c853c6e231"
#> [4] "bc05fa7d-fb0b-419b-9fc3-f3b5b0b21853"
#> [5] "848a9ce8-bc4d-446c-84ee-d77ccb5f08fd"
```
