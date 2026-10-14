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
#> [1] "1e68ef5c-ee95-40c6-86bf-7c23b58efe56"
#> [2] "4a00b521-3bbc-4942-8a8b-10d7c203e6ab"
#> [3] "292e4b31-82f1-4660-b942-3489f480759c"
#> [4] "c4d2ae48-a08d-437c-a6c3-746d788983d2"
#> [5] "4df16316-75b0-42cc-86fb-686ca905a1fa"
```
