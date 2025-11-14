# External Word document placeholder

Pour the content of a docx file in the resulting docx from an 'R
Markdown' document.

## Usage

``` r
block_pour_docx(file)
```

## Arguments

- file:

  external docx file path

## See also

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md),
[`block_section()`](https://davidgohel.github.io/officer/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/reference/unordered_list.md)

## Examples

``` r
library(officer)
docx <- tempfile(fileext = ".docx")
doc <- read_docx()
doc <- body_add(doc, iris[1:20,], style = "table_template")
print(doc, target = docx)

target <- tempfile(fileext = ".docx")
doc_1 <- read_docx()
doc_1 <- body_add(doc_1, block_pour_docx(docx))
print(doc_1, target = target)
```
