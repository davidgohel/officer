# Add character style in a Word document

The function lets you add or modify Word character styles.

## Usage

``` r
docx_set_character_style(
  x,
  style_id,
  style_name,
  base_on,
  fp_t = fp_text_lite()
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

  the character style name used as base style

- fp_t:

  Text formatting properties, see
  [`fp_text()`](https://davidgohel.github.io/officer/reference/fp_text.md).

## Examples

``` r
library(officer)
doc <- read_docx()

doc <- docx_set_character_style(
  doc,
  style_id = "newcharstyle",
  style_name = "label for char style",
  base_on = "Default Paragraph Font",
  fp_text_lite(
    shading.color = "red",
    color = "white")
)
paragraph <- fpar(
  run_wordtext("hello",
    style_id = "newcharstyle"))

doc <- body_add_fpar(doc, value = paragraph)
docx_file <- print(doc, target = tempfile(fileext = ".docx"))
docx_file
#> [1] "/tmp/RtmpWmrOW0/file178d4ab1fab2.docx"
```
