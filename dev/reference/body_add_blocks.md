# Add a list of blocks into a 'Word' document

Add a list of blocks produced by
[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md)
into into an rdocx object.

## Usage

``` r
body_add_blocks(x, blocks, pos = "after")
```

## Arguments

- x:

  an rdocx object

- blocks:

  set of blocks to be used as footnote content returned by function
  [`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md).

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

## See also

Other functions for adding content:
[`body_add_break()`](https://davidgohel.github.io/officer/dev/reference/body_add_break.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/dev/reference/body_add_caption.md),
[`body_add_docx()`](https://davidgohel.github.io/officer/dev/reference/body_add_docx.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/dev/reference/body_add_fpar.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/dev/reference/body_add_gg.md),
[`body_add_img()`](https://davidgohel.github.io/officer/dev/reference/body_add_img.md),
[`body_add_par()`](https://davidgohel.github.io/officer/dev/reference/body_add_par.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/dev/reference/body_add_plot.md),
[`body_add_table()`](https://davidgohel.github.io/officer/dev/reference/body_add_table.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/dev/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md),
[`body_import_docx()`](https://davidgohel.github.io/officer/dev/reference/body_import_docx.md)

## Examples

``` r
library(officer)

img.file <- file.path(R.home("doc"), "html", "logo.jpg")

bl <- block_list(
  fpar(ftext("hello", shortcuts$fp_bold(color = "red"))),
  fpar(
    ftext("hello world", shortcuts$fp_bold()),
    external_img(src = img.file, height = 1.06, width = 1.39),
    fp_p = fp_par(text.align = "center")
  )
)

doc_1 <- read_docx()
doc_1 <- body_add_blocks(doc_1, blocks = bl)
print(doc_1, target = tempfile(fileext = ".docx"))
```
