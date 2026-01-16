# Location for a placeholder based on a template

The function will return a list that complies with expected format for
argument `location` of function
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md).
A placeholder will be used as template and its positions will be updated
with values `left`, `top`, `width`, `height`.

## Usage

``` r
ph_location_template(
  left = 1,
  top = 1,
  width = 4,
  height = 3,
  newlabel = "",
  type = NULL,
  id = 1,
  ...
)
```

## Arguments

- left, top, width, height:

  place holder coordinates in inches.

- newlabel:

  a label for the placeholder. See section details.

- type:

  placeholder type to look for in the slide layout, one of 'body',
  'title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum'. It will be
  used as a template placeholder.

- id:

  index of the placeholder template. If two body placeholder, there can
  be two different index: 1 and 2 for the first and second body
  placeholders defined in the layout.

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
[`ph_location()`](https://davidgohel.github.io/officer/reference/ph_location.md),
[`ph_location_fullsize()`](https://davidgohel.github.io/officer/reference/ph_location_fullsize.md),
[`ph_location_id()`](https://davidgohel.github.io/officer/reference/ph_location_id.md),
[`ph_location_label()`](https://davidgohel.github.io/officer/reference/ph_location_label.md),
[`ph_location_left()`](https://davidgohel.github.io/officer/reference/ph_location_left.md),
[`ph_location_right()`](https://davidgohel.github.io/officer/reference/ph_location_right.md),
[`ph_location_type()`](https://davidgohel.github.io/officer/reference/ph_location_type.md)

## Examples

``` r
library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, "Title", location = ph_location_type(type = "title"))
doc <- ph_with(
  doc,
  "Hello world",
  location = ph_location_template(top = 4, type = "title")
)
print(doc, target = tempfile(fileext = ".pptx"))
```
