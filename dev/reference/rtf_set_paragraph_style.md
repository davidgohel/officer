# Add or replace a paragraph style in an RTF document

Add or replace a paragraph style in the document stylesheet so that
subsequent paragraphs can reference it via the `style` argument of
[`rtf_add()`](https://davidgohel.github.io/officer/dev/reference/rtf_add.md).
The function mirrors
[`docx_set_paragraph_style()`](https://davidgohel.github.io/officer/dev/reference/docx_set_paragraph_style.md)
for Word.

## Usage

``` r
rtf_set_paragraph_style(
  x,
  style_id,
  style_name = style_id,
  base_on = "Normal",
  fp_p = fp_par(),
  fp_t = NULL,
  outline_level = NULL
)
```

## Arguments

- x:

  an rtf object created by
  [`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md).

- style_id:

  user-facing identifier used as the key when a paragraph references the
  style (`rtf_add(..., style = style_id)`).

- style_name:

  display label of the style as it appears in Word's style menu.
  Defaults to `style_id`.

- base_on:

  the `style_id` of the parent style. Properties not defined in the new
  style are inherited from the parent. Defaults to `"Normal"`.

- fp_p:

  paragraph formatting properties, see
  [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md).

- fp_t:

  text formatting properties, see
  [`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md).
  If `NULL` the paragraph inherits text properties from the parent
  style.

- outline_level:

  integer between 1 and 9, or `NULL`. Sets the outline level
  (`\\outlinelevel`) so paragraphs using the style feed Word's
  navigation pane and TOC field. Level 1 is the topmost (used by Heading
  1).

## Value

the rtf object with the style added or updated.

## See also

[`docx_set_paragraph_style()`](https://davidgohel.github.io/officer/dev/reference/docx_set_paragraph_style.md),
[`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md),
[`rtf_add()`](https://davidgohel.github.io/officer/dev/reference/rtf_add.md)

## Examples

``` r
doc <- rtf_doc()
doc <- rtf_set_paragraph_style(
  doc,
  style_id = "Callout",
  fp_p = fp_par(text.align = "center", padding = 6),
  fp_t = fp_text_lite(bold = TRUE, color = "#1F4E79")
)
doc <- rtf_add(doc, "Heads up", style = "Callout")
print(doc, target = tempfile(fileext = ".rtf"))
#> [1] "/tmp/RtmpGthntp/file178b6c39ae99.rtf"
```
