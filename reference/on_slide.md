# Change current slide

Change current slide index of an rpptx object.

## Usage

``` r
on_slide(x, index)
```

## Arguments

- x:

  an rpptx object

- index:

  slide index

## See also

[`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md),
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md)

Other functions to manipulate slides:
[`add_slide()`](https://davidgohel.github.io/officer/reference/add_slide.md),
[`move_slide()`](https://davidgohel.github.io/officer/reference/move_slide.md),
[`remove_slide()`](https://davidgohel.github.io/officer/reference/remove_slide.md),
[`set_notes()`](https://davidgohel.github.io/officer/reference/set_notes.md)

## Examples

``` r
library(officer)

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- add_slide(doc, "Title and Content")
doc <- add_slide(doc, "Title and Content")
doc <- on_slide(doc, index = 1)
doc <- ph_with(
  x = doc,
  "First title",
  location = ph_location_type(type = "title")
)
doc <- on_slide(doc, index = 3)
doc <- ph_with(
  x = doc,
  "Third title",
  location = ph_location_type(type = "title")
)

file <- tempfile(fileext = ".pptx")
print(doc, target = file)
```
