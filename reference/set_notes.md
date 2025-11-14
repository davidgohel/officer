# Set notes for current slide

Set speaker notes for the current slide in a pptx presentation.

## Usage

``` r
set_notes(x, value, location, ...)

# S3 method for class 'character'
set_notes(x, value, location, ...)

# S3 method for class 'block_list'
set_notes(x, value, location, ...)
```

## Arguments

- x:

  an rpptx object

- value:

  text to be added to notes

- location:

  a placeholder location object. It will be used to specify the location
  of the new shape. This location can be defined with a call to one of
  the notes_ph functions. See section "see also".

- ...:

  further arguments passed to or from other methods.

## Methods (by class)

- `set_notes(character)`: add a character vector to a place holder in
  the notes on the current slide, values will be added as paragraphs.

- `set_notes(block_list)`: add a
  [`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md)
  to a place holder in the notes on the current slide.

## See also

[`print.rpptx()`](https://davidgohel.github.io/officer/reference/print.rpptx.md),
[`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md),
[`add_slide()`](https://davidgohel.github.io/officer/reference/add_slide.md),
[`notes_location_label()`](https://davidgohel.github.io/officer/reference/notes_location_label.md),
[`notes_location_type()`](https://davidgohel.github.io/officer/reference/notes_location_type.md)

Other functions to manipulate slides:
[`add_slide()`](https://davidgohel.github.io/officer/reference/add_slide.md),
[`move_slide()`](https://davidgohel.github.io/officer/reference/move_slide.md),
[`on_slide()`](https://davidgohel.github.io/officer/reference/on_slide.md),
[`remove_slide()`](https://davidgohel.github.io/officer/reference/remove_slide.md)

## Examples

``` r
# this name will be used to print the file
# change it to "youfile.pptx" to write the pptx
# file in your working directory.
fileout <- tempfile(fileext = ".pptx")
fpt_blue_bold <- fp_text_lite(color = "#006699", bold = TRUE)
doc <- read_pptx()
# add a slide with some text ----
doc <- add_slide(doc, layout = "Title and Content")
doc <- ph_with(x = doc, value = "Slide Title 1",
   location = ph_location_type(type = "title") )
# set speaker notes for the slide ----
doc <- set_notes(doc, value = "This text will only be visible for the speaker.",
   location = notes_location_type("body"))

# add a slide with some text ----
doc <- add_slide(doc, layout = "Title and Content")
doc <- ph_with(x = doc, value = "Slide Title 2",
   location = ph_location_type(type = "title") )
bl <- block_list(
  fpar(ftext("hello world", fpt_blue_bold)),
  fpar(ftext("Turlututu chapeau pointu", fpt_blue_bold))
)
doc <- set_notes(doc, value = bl,
   location = notes_location_type("body"))

print(doc, target = fileout)
```
