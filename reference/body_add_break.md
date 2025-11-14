# Add a page break in a 'Word' document

Add a page break into an rdocx object

## Usage

``` r
body_add_break(x, pos = "after")
```

## Arguments

- x:

  an rdocx object

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

## See also

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/reference/body_add_blocks.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/reference/body_add_caption.md),
[`body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/reference/body_add_fpar.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/reference/body_add_gg.md),
[`body_add_img()`](https://davidgohel.github.io/officer/reference/body_add_img.md),
[`body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/reference/body_add_plot.md),
[`body_add_table()`](https://davidgohel.github.io/officer/reference/body_add_table.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/reference/body_append_context.md),
[`body_import_docx()`](https://davidgohel.github.io/officer/reference/body_import_docx.md)

## Examples

``` r
doc <- read_docx()
doc <- body_add_break(doc)
print(doc, target = tempfile(fileext = ".docx"))
```
