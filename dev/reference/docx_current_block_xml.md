# xml element on which cursor is

Get the current block element as xml. This function is not to be used by
end users, it has been implemented to allow other packages to work with
officer. If the document is empty, this block will be set to NULL.

## Usage

``` r
docx_current_block_xml(x)
```

## Arguments

- x:

  an rdocx object

## Examples

``` r
doc <- read_docx()
docx_current_block_xml(doc)
#> NULL
```
