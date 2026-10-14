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
#> [1] "311a4a56-63ca-4bdd-8cb2-0af45c8a0dd6"
#> [2] "488d8a26-0211-4bc3-bbea-598911388b5c"
#> [3] "4709c734-92b3-4698-a325-2c5fa54184fd"
#> [4] "c118aee0-ab9d-4bee-97fe-f3480d808420"
#> [5] "7758f408-9dae-4a05-9efe-bc3e5212021d"
```
