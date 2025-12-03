# Add a slide

Add a slide into a pptx presentation.

## Usage

``` r
add_slide(x, layout = NULL, master = NULL, ..., .dots = NULL)
```

## Arguments

- x:

  an `rpptx` object.

- layout:

  slide layout name to use. Can be ommited of a default layout is set
  via
  [`layout_default()`](https://davidgohel.github.io/officer/dev/reference/layout_default.md).

- master:

  master layout name where `layout` is located. Only required in case of
  several masters if layout is not unique.

- ...:

  Key-value pairs of the form `"short form location" = object` passed to
  [phs_with](https://davidgohel.github.io/officer/dev/reference/phs_with.md).
  See section `"Short forms"` in
  [phs_with](https://davidgohel.github.io/officer/dev/reference/phs_with.md)
  for details, available short forms and examples.

- .dots:

  List of key-value pairs of the form
  `list("short form location" = object)`. Alternative to `...`. See
  [phs_with](https://davidgohel.github.io/officer/dev/reference/phs_with.md)
  for details.

## See also

[`print.rpptx()`](https://davidgohel.github.io/officer/dev/reference/print.rpptx.md),
[`read_pptx()`](https://davidgohel.github.io/officer/dev/reference/read_pptx.md),
[`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md),
[`phs_with()`](https://davidgohel.github.io/officer/dev/reference/phs_with.md),
[`layout_default()`](https://davidgohel.github.io/officer/dev/reference/layout_default.md)

Other functions to manipulate slides:
[`move_slide()`](https://davidgohel.github.io/officer/dev/reference/move_slide.md),
[`on_slide()`](https://davidgohel.github.io/officer/dev/reference/on_slide.md),
[`remove_slide()`](https://davidgohel.github.io/officer/dev/reference/remove_slide.md),
[`set_notes()`](https://davidgohel.github.io/officer/dev/reference/set_notes.md)

## Examples

``` r
x <- read_pptx()

layout_summary(x) # available layouts
#>              layout       master
#> 1       Title Slide Office Theme
#> 2 Title and Content Office Theme
#> 3    Section Header Office Theme
#> 4       Two Content Office Theme
#> 5        Comparison Office Theme
#> 6        Title Only Office Theme
#> 7             Blank Office Theme

x <- add_slide(x, layout = "Two Content")

x <- layout_default(x, "Title Slide") # set default layout for `add_slide()`
x <- add_slide(x) # uses default layout

# use `...` to fill placeholders when adding slide
x <- add_slide(
  x,
  layout = "Two Content",
  `Title 1` = "A title",
  dt = "Jan. 26, 2025",
  `body[2]` = "Body 2",
  left = "Left side",
  `6` = "Footer"
)
```
