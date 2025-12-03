# Empty block for 'PowerPoint'

Create an empty object to include as an empty placeholder shape in a
presentation. This comes in handy when presentation are updated through
R, but a user still wants to add some comments in this new content.

Empty content also works with layout fields (slide number and date) to
preserve them: they are included on the slide and keep being updated by
PowerPoint, i.e. update to the when the slide number when the slide
moves in the deck, update to the date.

## Usage

``` r
empty_content()
```

## See also

[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md),
[`body_add_blocks()`](https://davidgohel.github.io/officer/dev/reference/body_add_blocks.md)

## Examples

``` r
fileout <- tempfile(fileext = ".pptx")
doc <- read_pptx()
doc <- add_slide(doc, layout = "Two Content",
  master = "Office Theme")
doc <- ph_with(x = doc, value = empty_content(),
 location = ph_location_type(type = "title") )

doc <- add_slide(doc, "Title and Content")
# add slide number as a computer field
doc <- ph_with(
  x = doc, value = empty_content(),
  location = ph_location_type(type = "sldNum"))

print(doc, target = fileout )
```
