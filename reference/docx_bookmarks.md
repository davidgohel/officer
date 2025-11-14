# List Word bookmarks

List bookmarks id that can be found in a 'Word' document.

## Usage

``` r
docx_bookmarks(x)
```

## Arguments

- x:

  an `rdocx` object

## See also

Other functions for Word document informations:
[`doc_properties()`](https://davidgohel.github.io/officer/reference/doc_properties.md),
[`docx_dim()`](https://davidgohel.github.io/officer/reference/docx_dim.md),
[`length.rdocx()`](https://davidgohel.github.io/officer/reference/length.rdocx.md),
[`set_doc_properties()`](https://davidgohel.github.io/officer/reference/set_doc_properties.md),
[`styles_info()`](https://davidgohel.github.io/officer/reference/styles_info.md)

## Examples

``` r
library(officer)

doc_1 <- read_docx()
doc_1 <- body_add_par(doc_1, "centered text", style = "centered")
doc_1 <- body_bookmark(doc_1, "text_to_replace_1")
doc_1 <- body_add_par(doc_1, "centered text", style = "centered")
doc_1 <- body_bookmark(doc_1, "text_to_replace_2")

docx_bookmarks(doc_1)
#> [1] "text_to_replace_1" "text_to_replace_2"

docx_bookmarks(read_docx())
#> character(0)
```
