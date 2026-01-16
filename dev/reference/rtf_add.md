# Add content into an RTF document

This function add 'officer' objects into an RTF document. Values are
added as new paragraphs. See section 'Methods (by class)' that list
supported objects.

## Usage

``` r
rtf_add(x, value, ...)

# S3 method for class 'block_section'
rtf_add(x, value, ...)

# S3 method for class 'character'
rtf_add(x, value, ...)

# S3 method for class 'factor'
rtf_add(x, value, ...)

# S3 method for class 'double'
rtf_add(x, value, formatter = formatC, ...)

# S3 method for class 'fpar'
rtf_add(x, value, ...)

# S3 method for class 'block_list'
rtf_add(x, value, ...)

# S3 method for class 'gg'
rtf_add(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
)

# S3 method for class 'plot_instr'
rtf_add(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
)
```

## Arguments

- x:

  rtf object, created by
  [`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md).

- value:

  object to add in the document. Supported objects are vectors,
  graphics, block of formatted paragraphs. Use package 'flextable' to
  add tables.

- ...:

  further arguments passed to or from other methods. When adding a
  `ggplot` object or
  [plot_instr](https://davidgohel.github.io/officer/dev/reference/plot_instr.md),
  these arguments will be used by png function. See section 'Methods' to
  see what arguments can be used.

- formatter:

  function used to format the numerical values

- width:

  height in inches

- height:

  height in inches

- res:

  resolution of the png image in ppi

- scale:

  Multiplicative scaling factor, same as in ggsave

- ppr:

  [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)
  to apply to paragraph.

## Methods (by class)

- `rtf_add(block_section)`: add a new section definition

- `rtf_add(character)`: add characters as new paragraphs

- `rtf_add(factor)`: add a factor vector as new paragraphs

- `rtf_add(double)`: add a double vector as new paragraphs

- `rtf_add(fpar)`: add an
  [`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md)

- `rtf_add(block_list)`: add an
  [`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md)

- `rtf_add(gg)`: add a ggplot2

- `rtf_add(plot_instr)`: add a
  [`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md)
  object

## Examples

``` r
library(officer)

def_text <- fp_text_lite(color = "#006699", bold = TRUE)
center_par <- fp_par(text.align = "center", padding = 3)

doc <- rtf_doc(
  normal_par = fp_par(line_spacing = 1.4, padding = 3)
)

doc <- rtf_add(
  x = doc,
  value = fpar(
    ftext("how are you?", prop = def_text),
    fp_p = fp_par(text.align = "center")
  )
)

a_paragraph <- fpar(
  ftext("Here is a date: ", prop = def_text),
  run_word_field(field = "Date \\@ \"MMMM d yyyy\""),
  fp_p = center_par
)
doc <- rtf_add(
  x = doc,
  value = block_list(
    a_paragraph,
    a_paragraph,
    a_paragraph
  )
)

if (require("ggplot2")) {
  gg <- gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))
  doc <- rtf_add(doc, gg,
    width = 3, height = 4,
    ppr = center_par
  )
}
anyplot <- plot_instr(code = {
  barplot(1:5, col = 2:6)
})
doc <- rtf_add(doc, anyplot,
  width = 5, height = 4,
  ppr = center_par
)

print(doc, target = tempfile(fileext = ".rtf"))
#> [1] "/tmp/RtmpFsafPq/file1798272aa709.rtf"
```
