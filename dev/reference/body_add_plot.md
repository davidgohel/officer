# Add plot in a 'Word' document

Add a plot as a png image into an rdocx object.

## Usage

``` r
body_add_plot(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  style = "Normal",
  pos = "after",
  unit = "in",
  ...
)
```

## Arguments

- x:

  an rdocx object

- value:

  plot instructions, see
  [`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md).

- width, height:

  plot size in units expressed by the unit argument. Defaults to a width
  of 6 and a height of 5 "in"ches.

- res:

  resolution of the png image in ppi

- style:

  paragraph style

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

- unit:

  One of the following units in which the width and height arguments are
  expressed: "in", "cm" or "mm".

- ...:

  Arguments to be passed to png function.

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
[`body_add_table()`](https://davidgohel.github.io/officer/dev/reference/body_add_table.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/dev/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md),
[`body_import_docx()`](https://davidgohel.github.io/officer/dev/reference/body_import_docx.md)

## Examples

``` r
doc <- read_docx()

if (capabilities(what = "png")) {
  p <- plot_instr(
      code = {
        barplot(1:5, col = 2:6)
      }
    )

  doc <- body_add_plot(doc, value = p, style = "centered")

  # Set the unit in which the width and height arguments are expressed
  doc <- body_add_plot(doc, value = p, style = "centered", unit = "cm")
}

print(doc, target = tempfile(fileext = ".docx"))
```
