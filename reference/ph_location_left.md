# Location of a left body element

The function will return the location corresponding to a left bounding
box. The function assume the layout 'Two Content' is existing. This is
an helper function, if you don't have a layout named 'Two Content', use
[`ph_location_type()`](https://davidgohel.github.io/officer/reference/ph_location_type.md)
and set arguments to your specific needs.

## Usage

``` r
ph_location_left(newlabel = NULL, ...)
```

## Arguments

- newlabel:

  a label to associate with the placeholder.

- ...:

  unused arguments

## See also

Other functions for placeholder location:
[`ph_location()`](https://davidgohel.github.io/officer/reference/ph_location.md),
[`ph_location_fullsize()`](https://davidgohel.github.io/officer/reference/ph_location_fullsize.md),
[`ph_location_id()`](https://davidgohel.github.io/officer/reference/ph_location_id.md),
[`ph_location_label()`](https://davidgohel.github.io/officer/reference/ph_location_label.md),
[`ph_location_right()`](https://davidgohel.github.io/officer/reference/ph_location_right.md),
[`ph_location_template()`](https://davidgohel.github.io/officer/reference/ph_location_template.md),
[`ph_location_type()`](https://davidgohel.github.io/officer/reference/ph_location_type.md)

## Examples

``` r
library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, "Hello left", location = ph_location_left())
doc <- ph_with(doc, "Hello right", location = ph_location_right())
print(doc, target = tempfile(fileext = ".pptx"))
```
