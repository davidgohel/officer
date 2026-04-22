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
#> [1] "20df7a4d-d139-4f21-8705-6a15ac8d9610"
#> [2] "111cd352-2bb2-461b-9e8b-5645bdb23762"
#> [3] "5650af59-02da-4552-9d5b-9a9727b5ed32"
#> [4] "d17bd410-5f6b-4b62-ade0-c70919f9d7db"
#> [5] "667445e3-6c66-459d-bd09-5bcc3d71bb21"
```
