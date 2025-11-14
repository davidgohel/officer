# Add an image in a 'Word' document

Add an image into an rdocx object.

## Usage

``` r
body_add_img(x, src, style = NULL, width, height, pos = "after", unit = "in")
```

## Arguments

- x:

  an rdocx object

- src:

  image filename, the basename of the file must not contain any blank.

- style:

  paragraph style

- width, height:

  image size in units expressed by the unit argument. Defaults to
  "in"ches.

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

- unit:

  One of the following units in which the width and height arguments are
  expressed: "in", "cm" or "mm".

## See also

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/reference/body_add_blocks.md),
[`body_add_break()`](https://davidgohel.github.io/officer/reference/body_add_break.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/reference/body_add_caption.md),
[`body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/reference/body_add_fpar.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/reference/body_add_gg.md),
[`body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/reference/body_add_plot.md),
[`body_add_table()`](https://davidgohel.github.io/officer/reference/body_add_table.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/reference/body_append_context.md),
[`body_import_docx()`](https://davidgohel.github.io/officer/reference/body_import_docx.md)

## Examples

``` r
doc <- read_docx()

img.file <- file.path(R.home("doc"), "html", "logo.jpg")
if (file.exists(img.file)) {
  doc <- body_add_img(x = doc, src = img.file, height = 1.06, width = 1.39)

  # Set the unit in which the width and height arguments are expressed
  doc <- body_add_img(
    x = doc, src = img.file,
    height = 2.69, width = 3.53,
    unit = "cm"
  )
}

print(doc, target = tempfile(fileext = ".docx"))
```
