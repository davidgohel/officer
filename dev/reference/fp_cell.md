# Cell formatting properties

Create a `fp_cell` object that describes cell formatting properties.

## Usage

``` r
fp_cell(
  border = fp_border(width = 0),
  border.bottom,
  border.left,
  border.top,
  border.right,
  vertical.align = "center",
  margin = 0,
  margin.bottom,
  margin.top,
  margin.left,
  margin.right,
  background.color = "transparent",
  text.direction = "lrtb",
  rowspan = 1,
  colspan = 1
)

# S3 method for class 'fp_cell'
format(x, type = "wml", ...)

# S3 method for class 'fp_cell'
print(x, ...)

# S3 method for class 'fp_cell'
update(
  object,
  border,
  border.bottom,
  border.left,
  border.top,
  border.right,
  vertical.align,
  margin = 0,
  margin.bottom,
  margin.top,
  margin.left,
  margin.right,
  background.color,
  text.direction,
  rowspan = 1,
  colspan = 1,
  ...
)
```

## Arguments

- border:

  shortcut for all borders.

- border.bottom, border.left, border.top, border.right:

  [`fp_border()`](https://davidgohel.github.io/officer/dev/reference/fp_border.md)
  for borders.

- vertical.align:

  cell content vertical alignment - a single character value, expected
  value is one of "center" or "top" or "bottom"

- margin:

  shortcut for all margins.

- margin.bottom, margin.top, margin.left, margin.right:

  cell margins - 0 or positive integer value.

- background.color:

  cell background color - a single character value specifying a valid
  color (e.g. "#000000" or "black").

- text.direction:

  cell text rotation - a single character value, expected value is one
  of "lrtb", "tbrl", "btlr".

- rowspan:

  specify how many rows the cell is spanned over

- colspan:

  specify how many columns the cell is spanned over

- x, object:

  `fp_cell` object

- type:

  output type - one of 'wml', 'pml', 'html', 'rtf'.

- ...:

  further arguments - not used

## See also

Other functions for defining formatting properties:
[`fp_border()`](https://davidgohel.github.io/officer/dev/reference/fp_border.md),
[`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md),
[`fp_tab()`](https://davidgohel.github.io/officer/dev/reference/fp_tab.md),
[`fp_tabs()`](https://davidgohel.github.io/officer/dev/reference/fp_tabs.md),
[`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)

## Examples

``` r
obj <- fp_cell(margin = 1)
update(obj, margin.bottom = 5)
#> background-color:transparent;vertical-align:middle;border-bottom: 0.00pt solid transparent;border-top: 0.00pt solid transparent;border-left: 0.00pt solid transparent;border-right: 0.00pt solid transparent;margin-bottom:5pt;margin-top:1pt;margin-left:1pt;margin-right:1pt;
```
