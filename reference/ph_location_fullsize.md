# Location of a full size element

The function will return the location corresponding to a full size
display.

## Usage

``` r
ph_location_fullsize(newlabel = "", ...)
```

## Arguments

- newlabel:

  a label to associate with the placeholder.

- ...:

  unused arguments

## See also

Other functions for placeholder location:
[`ph_location()`](https://davidgohel.github.io/officer/reference/ph_location.md),
[`ph_location_id()`](https://davidgohel.github.io/officer/reference/ph_location_id.md),
[`ph_location_label()`](https://davidgohel.github.io/officer/reference/ph_location_label.md),
[`ph_location_left()`](https://davidgohel.github.io/officer/reference/ph_location_left.md),
[`ph_location_right()`](https://davidgohel.github.io/officer/reference/ph_location_right.md),
[`ph_location_template()`](https://davidgohel.github.io/officer/reference/ph_location_template.md),
[`ph_location_type()`](https://davidgohel.github.io/officer/reference/ph_location_type.md)

## Examples

``` r
library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, "Hello world", location = ph_location_fullsize())
print(doc, target = tempfile(fileext = ".pptx"))
```
