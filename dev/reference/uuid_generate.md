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
#> [1] "9b65fc20-43d2-4d5e-9aec-8fc73a6453de"
#> [2] "b8b8cf5c-e9c3-4798-8079-6c96b01dd6d2"
#> [3] "fd7d17d9-01e3-4106-a060-6b5e73ead8af"
#> [4] "70abc1d1-a726-4f08-8892-88f52394acc4"
#> [5] "7f19e741-ec11-474b-a6d5-6dcdacb4d1a0"
```
