# Write an 'RTF' File

Write the RTF object and its content to a file.

## Usage

``` r
# S3 method for class 'rtf'
print(x, target = NULL, ...)
```

## Arguments

- x:

  an 'rtf' object created with
  [`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md)

- target:

  path to the RTF file to write

- ...:

  unused

## See also

[`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md)

## Examples

``` r
# write a rdocx object in a rtf file ----
doc <- rtf_doc()
print(doc, target = tempfile(fileext = ".rtf"))
#> [1] "/tmp/RtmpxsgKBj/file17d85313bd50.rtf"
```
