# Slide layout properties plot

Plot slide layout properties into corresponding placeholders. This can
be useful to help visualize placeholders locations and identifiers.
*All* information in the plot stems from the
[`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md)
output. See *Details* section for more info.

## Usage

``` r
plot_layout_properties(
  x,
  layout = NULL,
  master = NULL,
  slide_idx = NULL,
  labels = TRUE,
  title = TRUE,
  type = TRUE,
  id = TRUE,
  cex = c(labels = 0.5, type = 0.5, id = 0.5),
  legend = FALSE
)
```

## Arguments

- x:

  an `rpptx` object

- layout:

  slide layout name or numeric index (row index from
  [`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md).
  If `NULL` (default), it plots the current slide's layout or the
  default layout (if set and there are not slides yet).

- master:

  master layout name where `layout` is located. Can be omitted if layout
  is unambiguous.

- slide_idx:

  Numeric slide index (default `NULL`) to specify which slide’s layout
  should be plotted.

- labels:

  if `TRUE` (default), adds placeholder labels (centered in *red*).

- title:

  if `TRUE` (default), adds a title with the layout and master name
  (latter in square brackets) at the top.

- type:

  if `TRUE` (default), adds the placeholder type and its index (in
  square brackets) in the upper left corner (in *blue*).

- id:

  if `TRUE` (default), adds the placeholder's unique `id` (see column
  `id` from
  [`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md))
  in the upper right corner (in *green*).

- cex:

  List or vector to specify font size for `labels`, `type`, and `id`.
  Default is `c(labels = .5, type = .5, id = .5)`. See
  [`graphics::text()`](https://rdrr.io/r/graphics/text.html) for details
  on how `cex` works. Matching by position and partial name matching is
  supported. A single numeric value will apply to all three parameters.

- legend:

  Add a legend to the plot (default `FALSE`).

## Details

The plot contains all relevant information to reference a placeholder
via the `ph_location_*` function family:

- `label`: ph label (red, center) to be used in
  [`ph_location_label()`](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md).
  *NB*: The label can be assigned by the user in PowerPoint.

- `type[idx]`: ph type + type index in brackets (blue, upper left) to be
  used in
  [`ph_location_type()`](https://davidgohel.github.io/officer/dev/reference/ph_location_type.md).
  *NB*: The index is consecutive and is sorted by ph position (top -\>
  bottom, left -\> right).

- `id`: ph id (green, upper right) to be used in
  [`ph_location_id()`](https://davidgohel.github.io/officer/dev/reference/ph_location_id.md)
  (forthcoming). *NB*: The id is set by PowerPoint automatically and
  lack a meaningful order.

## See also

Other functions for reading presentation information:
[`annotate_base()`](https://davidgohel.github.io/officer/dev/reference/annotate_base.md),
[`color_scheme()`](https://davidgohel.github.io/officer/dev/reference/color_scheme.md),
[`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md),
[`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md),
[`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md),
[`length.rpptx()`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md),
[`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md),
[`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)

## Examples

``` r
x <- read_pptx()

# select layout explicitly
plot_layout_properties(x = x, layout = "Title Slide", master = "Office Theme")

plot_layout_properties(x = x, layout = "Title Slide") # no master needed if layout name unique
plot_layout_properties(x = x, layout = 1) # use layout index instead of name

# plot default layout if one is set
x <- layout_default(x, "Title and Content")
plot_layout_properties(x)
#> ℹ Showing default layout: "Title and Content"


# plot current slide's layout (default if no layout is passed)
x <- add_slide(x, "Title Slide")
plot_layout_properties(x)
#> ℹ Showing current slide's layout: "Title Slide"


# specify which slide's layout to plot by index
plot_layout_properties(x, slide_idx = 1)
#> ℹ Showing layout of slide 1: "Title Slide"

# change appearance: what to show, font size, legend etc.
plot_layout_properties(
  x,
  layout = "Two Content",
  title = FALSE,
  type = FALSE,
  id = FALSE
)

plot_layout_properties(x, layout = 4, cex = c(labels = .8, id = .7, type = .7))

plot_layout_properties(x, 1, legend = TRUE)
```
