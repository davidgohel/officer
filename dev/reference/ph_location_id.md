# Location of a placeholder based on its id

Each placeholder has an id (a low integer value). The ids are unique
across a single layout. The function uses the placeholder's id to
reference it. Different from a ph label, the id is auto-assigned by
PowerPoint and cannot be modified by the user. Use
[`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md)
(column `id`) and
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md)
(upper right corner, in green) to find a placeholder's id.

## Usage

``` r
ph_location_id(id, newlabel = NULL, ...)
```

## Arguments

- id:

  placeholder id.

- newlabel:

  a new label to associate with the placeholder.

- ...:

  not used.

## Details

The location of the bounding box associated to a placeholder within a
slide is specified with the left top coordinate, the width and the
height. These are defined in inches:

- left:

  left coordinate of the bounding box

- top:

  top coordinate of the bounding box

- width:

  width of the bounding box

- height:

  height of the bounding box

In addition to these attributes, a label can be associated with the
shape. Shapes, text boxes, images and other objects will be identified
with that label in the *Selection Pane* of PowerPoint. This label can
then be reused by other functions such as
[`ph_location_label()`](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md).
It can be set with argument `newlabel`.

## See also

Other functions for placeholder location:
[`ph_location()`](https://davidgohel.github.io/officer/dev/reference/ph_location.md),
[`ph_location_fullsize()`](https://davidgohel.github.io/officer/dev/reference/ph_location_fullsize.md),
[`ph_location_label()`](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md),
[`ph_location_left()`](https://davidgohel.github.io/officer/dev/reference/ph_location_left.md),
[`ph_location_right()`](https://davidgohel.github.io/officer/dev/reference/ph_location_right.md),
[`ph_location_template()`](https://davidgohel.github.io/officer/dev/reference/ph_location_template.md),
[`ph_location_type()`](https://davidgohel.github.io/officer/dev/reference/ph_location_type.md)

## Examples

``` r
doc <- read_pptx()
doc <- add_slide(doc, "Comparison")
plot_layout_properties(doc, "Comparison")


doc <- ph_with(doc, "The Title", location = ph_location_id(id = 2)) # title
doc <- ph_with(doc, "Left Header", location = ph_location_id(id = 3)) # left header
doc <- ph_with(doc, "Left Content", location = ph_location_id(id = 4)) # left content
doc <- ph_with(doc, "The Footer", location = ph_location_id(id = 8)) # footer

file <- tempfile(fileext = ".pptx")
print(doc, file)
if (FALSE) { # \dontrun{
file.show(file) # may not work on your system
} # }
```
