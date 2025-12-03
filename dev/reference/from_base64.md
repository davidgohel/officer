# Decode Base64 Vector to Character

Decodes one or more Base64-encoded elements back into their original
character form.

## Usage

``` r
from_base64(x)
```

## Arguments

- x:

  A character vector of Base64-encoded strings. NA values are preserved.

## Value

A character vector of decoded strings, same length as `x`.

## Examples

``` r
z <- as_base64(c("hello", "world"))
from_base64(z)
#> [1] "hello" "world"
```
