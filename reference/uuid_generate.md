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
#> [1] "33334928-9d0d-4397-935c-24317ec92f7d"
#> [2] "a0720276-ada4-4329-a412-15ff766c2d71"
#> [3] "8ed30d48-1864-41d1-bd9b-af760514c6ad"
#> [4] "82fbb105-d4bc-4b43-a709-7b84f3d189dd"
#> [5] "b3bcea46-0a57-4315-9142-de83b4ec18e1"
```
