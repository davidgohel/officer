# Add fpar in a 'Word' document

Add an
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md) (a
formatted paragraph) into an rdocx object.

## Usage

``` r
body_add_fpar(x, value, style = NULL, pos = "after")
```

## Arguments

- x:

  a docx device

- value:

  a character

- style:

  paragraph style. If NULL, paragraph settings from `fpar` will be used.
  If not NULL, it must be a paragraph style name (located in the
  template provided as `read_docx(path = ...)`); in that case, paragraph
  settings from `fpar` will be ignored.

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

## See also

[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md)

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/reference/body_add_blocks.md),
[`body_add_break()`](https://davidgohel.github.io/officer/reference/body_add_break.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/reference/body_add_caption.md),
[`body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.md),
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
bold_face <- shortcuts$fp_bold(font.size = 30)
bold_redface <- update(bold_face, color = "red")
fpar_ <- fpar(
  ftext("Hello ", prop = bold_face),
  ftext("World", prop = bold_redface),
  ftext(", how are you?", prop = bold_face)
)
doc <- read_docx()
doc <- body_add_fpar(doc, fpar_)

print(doc, target = tempfile(fileext = ".docx"))

# a way of using fpar to center an image in a Word doc ----
rlogo <- file.path(R.home("doc"), "html", "logo.jpg")
img_in_par <- fpar(
  external_img(src = rlogo, height = 1.06 / 2, width = 1.39 / 2),
  hyperlink_ftext(
    href = "https://cran.r-project.org/index.html",
    text = "cran", prop = bold_redface
  ),
  fp_p = fp_par(text.align = "center")
)

doc <- read_docx()
doc <- body_add_fpar(doc, img_in_par)
print(doc, target = tempfile(fileext = ".docx"))
```
