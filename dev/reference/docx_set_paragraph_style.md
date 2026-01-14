# Add or replace paragraph style in a Word document

The function lets you add or replace a Word paragraph style.

## Usage

``` r
docx_set_paragraph_style(
  x,
  style_id,
  style_name,
  base_on = "Normal",
  fp_p = fp_par(),
  fp_t = NULL
)
```

## Arguments

- x:

  an rdocx object

- style_id:

  a unique style identifier for Word.

- style_name:

  a unique label associated with the style identifier. This label is the
  name of the style when Word edit the document.

- base_on:

  the style name used as base style

- fp_p:

  paragraph formatting properties, see
  [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md).

- fp_t:

  default text formatting properties. This is used as text formatting
  properties, see
  [`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md).
  If NULL (default), the paragraph will used the default text formatting
  properties (defined by the `base_on` argument).

## Examples

``` r
library(officer)

doc <- read_docx()

doc <- docx_set_paragraph_style(
  doc,
  style_id = "rightaligned",
  style_name = "Explicit label",
  fp_p = fp_par(text.align = "right", padding = 20),
  fp_t = fp_text_lite(
    bold = TRUE,
    shading.color = "#FD34F0",
    color = "white")
)

doc <- body_add_par(doc,
  value = "This is a test",
  style = "Explicit label")

docx_file <- print(doc, target = tempfile(fileext = ".docx"))
docx_file
#> [1] "/tmp/Rtmp6xWNVi/file176a1cb62690.docx"
```
