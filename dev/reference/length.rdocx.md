# Number of blocks inside an rdocx object

Return the number of blocks inside an rdocx object. This number also
include the default section definition of a Word document - default Word
section is an uninvisible element.

## Usage

``` r
# S3 method for class 'rdocx'
length(x)
```

## Arguments

- x:

  an rdocx object

## See also

Other functions for Word document informations:
[`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md),
[`docx_bookmarks()`](https://davidgohel.github.io/officer/dev/reference/docx_bookmarks.md),
[`docx_dim()`](https://davidgohel.github.io/officer/dev/reference/docx_dim.md),
[`set_doc_properties()`](https://davidgohel.github.io/officer/dev/reference/set_doc_properties.md),
[`styles_info()`](https://davidgohel.github.io/officer/dev/reference/styles_info.md)

## Examples

``` r
# how many elements are there in an new document produced
# with the default template.
length( read_docx() )
#> [1] 1
```
