# Move a slide

Move a slide in a pptx presentation.

## Usage

``` r
move_slide(x, index = NULL, to)
```

## Arguments

- x:

  an rpptx object

- index:

  slide index, default to current slide position.

- to:

  new slide index.

## Note

cursor is set on the last slide.

## See also

[`read_pptx()`](https://davidgohel.github.io/officer/dev/reference/read_pptx.md)

Other functions to manipulate slides:
[`add_slide()`](https://davidgohel.github.io/officer/dev/reference/add_slide.md),
[`on_slide()`](https://davidgohel.github.io/officer/dev/reference/on_slide.md),
[`remove_slide()`](https://davidgohel.github.io/officer/dev/reference/remove_slide.md),
[`set_notes()`](https://davidgohel.github.io/officer/dev/reference/set_notes.md)

## Examples

``` r
x <- read_pptx()
x <- add_slide(x, "Title and Content")
x <- ph_with(x, "Hello world 1", location = ph_location_type())
x <- add_slide(x, "Title and Content")
x <- ph_with(x, "Hello world 2", location = ph_location_type())
x <- move_slide(x, index = 1, to = 2)
```
