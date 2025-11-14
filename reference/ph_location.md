# Location for a placeholder from scratch

The function will return a list that complies with expected format for
argument `location` of function
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md).

## Usage

``` r
ph_location(
  left = 1,
  top = 1,
  width = 4,
  height = 3,
  newlabel = "",
  bg = NULL,
  rotation = NULL,
  ln = NULL,
  geom = NULL,
  ...
)
```

## Arguments

- left, top, width, height:

  place holder coordinates in inches.

- newlabel:

  a label for the placeholder. See section details.

- bg:

  background color

- rotation:

  rotation angle

- ln:

  a
  [`sp_line()`](https://davidgohel.github.io/officer/reference/sp_line.md)
  object specifying the outline style.

- geom:

  shape geometry, see
  http://www.datypic.com/sc/ooxml/t-a_ST_ShapeType.html

- ...:

  unused arguments

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
[`ph_location_label()`](https://davidgohel.github.io/officer/reference/ph_location_label.md).
It can be set with argument `newlabel`.

## See also

Other functions for placeholder location:
[`ph_location_fullsize()`](https://davidgohel.github.io/officer/reference/ph_location_fullsize.md),
[`ph_location_id()`](https://davidgohel.github.io/officer/reference/ph_location_id.md),
[`ph_location_label()`](https://davidgohel.github.io/officer/reference/ph_location_label.md),
[`ph_location_left()`](https://davidgohel.github.io/officer/reference/ph_location_left.md),
[`ph_location_right()`](https://davidgohel.github.io/officer/reference/ph_location_right.md),
[`ph_location_template()`](https://davidgohel.github.io/officer/reference/ph_location_template.md),
[`ph_location_type()`](https://davidgohel.github.io/officer/reference/ph_location_type.md)

## Examples

``` r
doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, "Hello world",
  location = ph_location(width = 4, height = 3, newlabel = "hello")
)
print(doc, target = tempfile(fileext = ".pptx"))

# Set geometry and outline
doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
loc <- ph_location(
  left = 1, top = 1, width = 4, height = 3, bg = "steelblue",
  ln = sp_line(color = "red", lwd = 2.5),
  geom = "trapezoid"
)
doc <- ph_with(doc, "", loc = loc)
print(doc, target = tempfile(fileext = ".pptx"))
```
