# pptx tags for visual and non visual properties

Visual and non visual properties of a shape can be returned by this
function.

## Usage

``` r
shape_properties_tags(
  left = 0,
  top = 0,
  width = 3,
  height = 3,
  bg = "transparent",
  rot = 0,
  label = "",
  ph = "<p:ph/>",
  ln = sp_line(lwd = 0, linecmpd = "solid", lineend = "rnd"),
  geom = NULL
)
```

## Arguments

- left, top, width, height:

  place holder coordinates in inches.

- bg:

  background color

- rot:

  rotation angle

- label:

  a label for the placeholder.

- ph:

  string containing xml code for ph tag

- ln:

  a
  [`sp_line()`](https://davidgohel.github.io/officer/reference/sp_line.md)
  object specifying the outline style.

- geom:

  shape geometry, see
  http://www.datypic.com/sc/ooxml/t-a_ST_ShapeType.html

## Value

a character value

## See also

Other functions for officer extensions:
[`fortify_location()`](https://davidgohel.github.io/officer/reference/fortify_location.md),
[`get_reference_value()`](https://davidgohel.github.io/officer/reference/get_reference_value.md),
[`opts_current_table()`](https://davidgohel.github.io/officer/reference/opts_current_table.md),
[`str_encode_to_rtf()`](https://davidgohel.github.io/officer/reference/str_encode_to_rtf.md),
[`to_html()`](https://davidgohel.github.io/officer/reference/to_html.md),
[`to_pml()`](https://davidgohel.github.io/officer/reference/to_pml.md),
[`to_rtf()`](https://davidgohel.github.io/officer/reference/to_rtf.md),
[`to_wml()`](https://davidgohel.github.io/officer/reference/to_wml.md),
[`wml_link_images()`](https://davidgohel.github.io/officer/reference/wml_link_images.md)
