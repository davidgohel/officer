# Add Word caption in a 'Word' document

Add a Word caption into an rdocx object.

## Usage

``` r
body_add_caption(x, value, pos = "after")
```

## Arguments

- x:

  an rdocx object

- value:

  an object returned by
  [`block_caption()`](https://davidgohel.github.io/officer/dev/reference/block_caption.md)

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

## See also

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/dev/reference/body_add_blocks.md),
[`body_add_break()`](https://davidgohel.github.io/officer/dev/reference/body_add_break.md),
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
doc <- read_docx()

if (capabilities(what = "png")) {
  doc <- body_add_plot(doc,
    value = plot_instr(
      code = {
        barplot(1:5, col = 2:6)
      }
    ),
    style = "centered"
  )
}
run_num <- run_autonum(
  seq_id = "fig", pre_label = "Figure ",
  bkm = "barplot"
)
caption <- block_caption("a barplot",
  style = "Normal",
  autonum = run_num
)
doc <- body_add_caption(doc, caption)
print(doc, target = tempfile(fileext = ".docx"))
```
