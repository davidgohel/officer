# Wrap plot instructions for png plotting in Powerpoint or Word

A simple wrapper to capture plot instructions that will be executed and
copied in a document. It produces an object of class 'plot_instr' with a
corresponding method
[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
and
[`body_add_plot()`](https://davidgohel.github.io/officer/dev/reference/body_add_plot.md).

The function enable usage of any R plot with argument `code`. Wrap your
code between curly bracket if more than a single expression.

## Usage

``` r
plot_instr(code)
```

## Arguments

- code:

  plotting instructions

## See also

[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/dev/reference/body_add_plot.md)

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/dev/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/dev/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/dev/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/dev/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/dev/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/dev/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md),
[`unordered_list()`](https://davidgohel.github.io/officer/dev/reference/unordered_list.md)

## Examples

``` r
# plot_instr demo ----

anyplot <- plot_instr(code = {
  barplot(1:5, col = 2:6)
  })

doc <- read_docx()
doc <- body_add(doc, anyplot, width = 5, height = 4)
print(doc, target = tempfile(fileext = ".docx"))


doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(
  doc, anyplot,
  location = ph_location_fullsize(),
  bg = "#00000066", pointsize = 12)
print(doc, target = tempfile(fileext = ".pptx"))
```
