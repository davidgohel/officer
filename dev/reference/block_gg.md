# 'ggplot' block

A simple wrapper to add a 'ggplot' object as a png in a document. It
produces an object of class 'block_gg' with a corresponding method
[`to_wml()`](https://davidgohel.github.io/officer/dev/reference/to_wml.md)
that can be used to convert the object to a WordML string.

## Usage

``` r
block_gg(
  value,
  fp_p = fp_par(),
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  unit = "in"
)
```

## Arguments

- value:

  'ggplot' object

- fp_p:

  paragraph formatting properties, see
  [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)

- width, height:

  image size in units expressed by the unit argument. Defaults to
  "in"ches.

- res:

  resolution of the png image in ppi

- scale:

  Multiplicative scaling factor, same as in ggsave

- unit:

  One of the following units in which the width and height arguments are
  expressed: "in", "cm" or "mm".

## See also

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/dev/reference/block_caption.md),
[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/dev/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/dev/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/dev/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/dev/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/dev/reference/unordered_list.md)

## Examples

``` r
library(officer)
if (require("ggplot2")) {
  set.seed(2)
  doc <- read_docx()

  z <- body_append_start_context(doc)

  for (i in seq_len(3)) {
    df <- data.frame(x = runif(10), y = runif(10))
    gg <- ggplot(df, aes(x = x, y = y)) + geom_point()

    write_elements_to_context(
      context = z,
      block_gg(
        value = gg
      )
    )
  }

  doc <- body_append_stop_context(z)

  print(doc, target = tempfile(fileext = ".docx"))
}
#> Loading required package: ggplot2
```
