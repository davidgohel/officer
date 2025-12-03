# Add table in a 'Word' document

Add a table into an rdocx object.

## Usage

``` r
body_add_table(
  x,
  value,
  style = NULL,
  pos = "after",
  header = TRUE,
  alignment = NULL,
  align_table = "center",
  stylenames = table_stylenames(),
  first_row = TRUE,
  first_column = FALSE,
  last_row = FALSE,
  last_column = FALSE,
  no_hband = FALSE,
  no_vband = TRUE
)
```

## Arguments

- x:

  a docx device

- value:

  a data.frame to add as a table

- style:

  table style

- pos:

  where to add the new element relative to the cursor, one of after",
  "before", "on".

- header:

  display header if TRUE

- alignment:

  columns alignement, argument length must match with columns length,
  values must be "l" (left), "r" (right) or "c" (center).

- align_table:

  table alignment within document, value must be "left", "center" or
  "right"

- stylenames:

  columns styles defined by
  [`table_stylenames()`](https://davidgohel.github.io/officer/dev/reference/table_stylenames.md)

- first_row:

  Specifies that the first column conditional formatting should be
  applied. Details for this and other conditional formatting options can
  be found at http://officeopenxml.com/WPtblLook.php.

- first_column:

  Specifies that the first column conditional formatting should be
  applied.

- last_row:

  Specifies that the first column conditional formatting should be
  applied.

- last_column:

  Specifies that the first column conditional formatting should be
  applied.

- no_hband:

  Specifies that the first column conditional formatting should be
  applied.

- no_vband:

  Specifies that the first column conditional formatting should be
  applied.

## See also

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/dev/reference/body_add_blocks.md),
[`body_add_break()`](https://davidgohel.github.io/officer/dev/reference/body_add_break.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/dev/reference/body_add_caption.md),
[`body_add_docx()`](https://davidgohel.github.io/officer/dev/reference/body_add_docx.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/dev/reference/body_add_fpar.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/dev/reference/body_add_gg.md),
[`body_add_img()`](https://davidgohel.github.io/officer/dev/reference/body_add_img.md),
[`body_add_par()`](https://davidgohel.github.io/officer/dev/reference/body_add_par.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/dev/reference/body_add_plot.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/dev/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md),
[`body_import_docx()`](https://davidgohel.github.io/officer/dev/reference/body_import_docx.md)

## Examples

``` r
doc <- read_docx()
doc <- body_add_table(doc, iris, style = "table_template")

print(doc, target = tempfile(fileext = ".docx"))
```
