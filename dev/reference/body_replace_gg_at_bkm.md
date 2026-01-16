# Add plots at bookmark location in a 'Word' document

Use these functions if you want to replace a paragraph containing a
bookmark with a 'ggplot' or a base plot.

## Usage

``` r
body_replace_gg_at_bkm(
  x,
  bookmark,
  value,
  width = 6,
  height = 5,
  res = 300,
  style = "Normal",
  scale = 1,
  keep = FALSE,
  ...
)

body_replace_plot_at_bkm(
  x,
  bookmark,
  value,
  width = 6,
  height = 5,
  res = 300,
  style = "Normal",
  keep = FALSE,
  ...
)
```

## Arguments

- x:

  an rdocx object

- bookmark:

  bookmark id

- value:

  a ggplot object for body_replace_gg_at_bkm() or a set plot
  instructions body_replace_plot_at_bkm(), see plot_instr().

- width, height:

  plot size in units expressed by the unit argument. Defaults to a width
  of 6 and a height of 5 "in"ches.

- res:

  resolution of the png image in ppi

- style:

  paragraph style

- scale:

  Multiplicative scaling factor, same as in ggsave

- keep:

  Should the bookmark be preserved? Defaults to `FALSE`, i.e.the
  bookmark will be lost as a side effect.

- ...:

  Arguments to be passed to png function.
