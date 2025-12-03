# Add paragraphs of text in a 'Word' document

Add a paragraph of text into an rdocx object

## Usage

``` r
body_add_par(x, value, style = NULL, pos = "after")
```

## Arguments

- x:

  a docx device

- value:

  a character

- style:

  paragraph style name

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

## See also

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/dev/reference/body_add_blocks.md),
[`body_add_break()`](https://davidgohel.github.io/officer/dev/reference/body_add_break.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/dev/reference/body_add_caption.md),
[`body_add_docx()`](https://davidgohel.github.io/officer/dev/reference/body_add_docx.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/dev/reference/body_add_fpar.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/dev/reference/body_add_gg.md),
[`body_add_img()`](https://davidgohel.github.io/officer/dev/reference/body_add_img.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/dev/reference/body_add_plot.md),
[`body_add_table()`](https://davidgohel.github.io/officer/dev/reference/body_add_table.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/dev/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md),
[`body_import_docx()`](https://davidgohel.github.io/officer/dev/reference/body_import_docx.md)

## Examples

``` r
doc <- read_docx()
doc <- body_add_par(doc, "A title", style = "heading 1")
doc <- body_add_par(doc, "Hello world!", style = "Normal")
doc <- body_add_par(doc, "centered text", style = "centered")

print(doc, target = tempfile(fileext = ".docx"))
```
