# Location of a placeholder based on a type

The function will use the type name of the placeholder (e.g. body,
title), the layout name and few other criterias to find the
corresponding location.

## Usage

``` r
ph_location_type(
  type = "body",
  type_idx = NULL,
  position_right = TRUE,
  position_top = TRUE,
  newlabel = NULL,
  id = NULL,
  ...
)
```

## Arguments

- type:

  placeholder type to look for in the slide layout, one of 'body',
  'title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum'.

- type_idx:

  Type index of the placeholder. If there is more than one placeholder
  of a type (e.g., `body`), the type index can be supplied to uniquely
  identify a ph. The index is a running number starting at 1. It is
  assigned by placeholder position (top -\> bottom, left -\> right). See
  [`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md)
  for details. If `idx` argument is used, `position_right` and
  `position_top` are ignored.

- position_right:

  the parameter is used when a selection with above parameters does not
  provide a unique position (for example layout 'Two Content' contains
  two element of type 'body'). If `TRUE`, the element the most on the
  right side will be selected, otherwise the element the most on the
  left side will be selected.

- position_top:

  same than `position_right` but applied to top versus bottom.

- newlabel:

  a label to associate with the placeholder.

- id:

  (**DEPRECATED, use `type_idx` instead**) Index of the placeholder. If
  two body placeholder, there can be two different index: 1 and 2 for
  the first and second body placeholders defined in the layout. If this
  argument is used, `position_right` and `position_top` will be ignored.

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
[`ph_location_label()`](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md).
It can be set with argument `newlabel`.

## See also

Other functions for placeholder location:
[`ph_location()`](https://davidgohel.github.io/officer/dev/reference/ph_location.md),
[`ph_location_fullsize()`](https://davidgohel.github.io/officer/dev/reference/ph_location_fullsize.md),
[`ph_location_id()`](https://davidgohel.github.io/officer/dev/reference/ph_location_id.md),
[`ph_location_label()`](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md),
[`ph_location_left()`](https://davidgohel.github.io/officer/dev/reference/ph_location_left.md),
[`ph_location_right()`](https://davidgohel.github.io/officer/dev/reference/ph_location_right.md),
[`ph_location_template()`](https://davidgohel.github.io/officer/dev/reference/ph_location_template.md)

## Examples

``` r
# ph_location_type demo ----

loc_title <- ph_location_type(type = "title")
loc_footer <- ph_location_type(type = "ftr")
loc_dt <- ph_location_type(type = "dt")
loc_slidenum <- ph_location_type(type = "sldNum")
loc_body <- ph_location_type(type = "body")


doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(x = doc, "Un titre", location = loc_title)
doc <- ph_with(x = doc, "pied de page", location = loc_footer)
doc <- ph_with(x = doc, format(Sys.Date()), location = loc_dt)
doc <- ph_with(x = doc, "slide 1", location = loc_slidenum)
doc <- ph_with(x = doc, letters[1:10], location = loc_body)

loc_subtitle <- ph_location_type(type = "subTitle")
loc_ctrtitle <- ph_location_type(type = "ctrTitle")
doc <- add_slide(doc, layout = "Title Slide")
doc <- ph_with(x = doc, "Un sous titre", location = loc_subtitle)
doc <- ph_with(x = doc, "Un titre", location = loc_ctrtitle)

fileout <- tempfile(fileext = ".pptx")
print(doc, target = fileout)
```
