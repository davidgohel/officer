# Paragraph formatting properties

Create a `fp_par` object that describes paragraph formatting properties.

Function `fp_par_lite()` is generating properties with only entries for
the parameters users provided. The undefined properties will inherit
from the default settings.

## Usage

``` r
fp_par(
  text.align = "left",
  padding = 0,
  line_spacing = 1,
  border = fp_border(width = 0),
  padding.bottom,
  padding.top,
  padding.left,
  padding.right,
  border.bottom,
  border.left,
  border.top,
  border.right,
  shading.color = "transparent",
  keep_with_next = FALSE,
  tabs = NULL,
  word_style = "Normal"
)

fp_par_lite(
  text.align = NA,
  padding = NA,
  line_spacing = NA,
  border = FALSE,
  padding.bottom = NA,
  padding.top = NA,
  padding.left = NA,
  padding.right = NA,
  border.bottom = FALSE,
  border.left = FALSE,
  border.top = FALSE,
  border.right = FALSE,
  shading.color = NA,
  keep_with_next = NA,
  tabs = FALSE,
  word_style = NA
)

# S3 method for class 'fp_par'
print(x, ...)

# S3 method for class 'fp_par'
update(
  object,
  text.align,
  padding,
  border,
  padding.bottom,
  padding.top,
  padding.left,
  padding.right,
  border.bottom,
  border.left,
  border.top,
  border.right,
  shading.color,
  keep_with_next,
  word_style,
  ...
)
```

## Arguments

- text.align:

  text alignment - a single character value, expected value is one of
  'left', 'right', 'center', 'justify'.

- padding:

  paragraph paddings - 0 or positive integer value. Argument `padding`
  overwrites arguments `padding.bottom`, `padding.top`, `padding.left`,
  `padding.right`.

- line_spacing:

  line spacing, 1 is single line spacing, 2 is double line spacing.

- border:

  shortcut for all borders.

- padding.bottom, padding.top, padding.left, padding.right:

  paragraph paddings - 0 or positive integer value.

- border.bottom, border.left, border.top, border.right:

  [`fp_border()`](https://davidgohel.github.io/officer/dev/reference/fp_border.md)
  for borders. overwrite other border properties.

- shading.color:

  shading color - a single character value specifying a valid color
  (e.g. "#000000" or "black").

- keep_with_next:

  a scalar logical. Specifies that the paragraph (or at least part of
  it) should be rendered on the same page as the next paragraph when
  possible.

- tabs:

  NULL (default) for no tabulation marks setting or an object returned
  by
  [`fp_tabs()`](https://davidgohel.github.io/officer/dev/reference/fp_tabs.md).
  Note this can only have effect with Word or RTF outputs.

- word_style:

  Word paragraph style name

- x, object:

  `fp_par` object

- ...:

  further arguments - not used

## Value

a `fp_par` object

## See also

[fpar](https://davidgohel.github.io/officer/dev/reference/fpar.md)

Other functions for defining formatting properties:
[`fp_border()`](https://davidgohel.github.io/officer/dev/reference/fp_border.md),
[`fp_cell()`](https://davidgohel.github.io/officer/dev/reference/fp_cell.md),
[`fp_tab()`](https://davidgohel.github.io/officer/dev/reference/fp_tab.md),
[`fp_tabs()`](https://davidgohel.github.io/officer/dev/reference/fp_tabs.md),
[`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)

## Examples

``` r
fp_par(text.align = "center", padding = 5)
#>                     values
#> text.align          center
#> padding.top              5
#> padding.bottom           5
#> padding.left             5
#> padding.right            5
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
obj <- fp_par(text.align = "center", padding = 1)
update(obj, padding.bottom = 5)
#>                     values
#> text.align          center
#> padding.top              1
#> padding.bottom           5
#> padding.left             1
#> padding.right            1
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
```
