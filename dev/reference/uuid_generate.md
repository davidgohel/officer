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
#> [1] "db821e00-1807-4c77-8557-fc724a49f88c"
#> [2] "ec00d720-0863-4271-a19c-4b52b1ceaa7f"
#> [3] "bf643e34-c63b-4b11-886f-6ea3b5d51be7"
#> [4] "e3f111df-7d8f-47f0-aee8-ebc6d039bd59"
#> [5] "778e450a-84d4-467d-a061-21c6dd2dd98b"
```
