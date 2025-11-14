# Remove a slide

Remove a slide from a pptx presentation.

## Usage

``` r
remove_slide(x, index = NULL, rm_images = FALSE)
```

## Arguments

- x:

  an rpptx object

- index:

  slide index, default to current slide position.

- rm_images:

  unused anymore.

## Note

cursor is set on the last slide.

## See also

[`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md),
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md),
[`ph_remove()`](https://davidgohel.github.io/officer/reference/ph_remove.md)

Other functions to manipulate slides:
[`add_slide()`](https://davidgohel.github.io/officer/reference/add_slide.md),
[`move_slide()`](https://davidgohel.github.io/officer/reference/move_slide.md),
[`on_slide()`](https://davidgohel.github.io/officer/reference/on_slide.md),
[`set_notes()`](https://davidgohel.github.io/officer/reference/set_notes.md)

## Examples

``` r
my_pres <- read_pptx()
my_pres <- add_slide(my_pres, "Title and Content")
my_pres <- remove_slide(my_pres)
```
