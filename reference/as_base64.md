# Encode Character Vector to Base64

Encodes one or more elements of a character vector into Base64 format.

## Usage

``` r
as_base64(x)
```

## Arguments

- x:

  A character vector. NA values are preserved.

## Value

A character vector of Base64-encoded strings, same length as `x`.

## Examples

``` r
as_base64(letters)
#>  [1] "YQ==" "Yg==" "Yw==" "ZA==" "ZQ==" "Zg==" "Zw==" "aA==" "aQ==" "ag=="
#> [11] "aw==" "bA==" "bQ==" "bg==" "bw==" "cA==" "cQ==" "cg==" "cw==" "dA=="
#> [21] "dQ==" "dg==" "dw==" "eA==" "eQ==" "eg=="
as_base64(c("hello", NA, "world"))
#> [1] "aGVsbG8=" NA         "d29ybGQ="
```
