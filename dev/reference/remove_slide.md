# Remove slide(s)

Remove one or more slides from a pptx presentation.

## Usage

``` r
remove_slide(x, index = NULL, rm_images = FALSE)
```

## Arguments

- x:

  an rpptx object

- index:

  slide index or a vector of slide indices to remove, default to current
  slide position.

- rm_images:

  unused anymore.

## Note

cursor is set on the last slide.

## See also

[`read_pptx()`](https://davidgohel.github.io/officer/dev/reference/read_pptx.md),
[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md),
[`ph_remove()`](https://davidgohel.github.io/officer/dev/reference/ph_remove.md)

Other functions to manipulate slides:
[`add_slide()`](https://davidgohel.github.io/officer/dev/reference/add_slide.md),
[`move_slide()`](https://davidgohel.github.io/officer/dev/reference/move_slide.md),
[`on_slide()`](https://davidgohel.github.io/officer/dev/reference/on_slide.md),
[`set_notes()`](https://davidgohel.github.io/officer/dev/reference/set_notes.md)

## Examples

``` r
library(officer)

x <- read_pptx()
x <- add_slide(x, "Title and Content")
x <- remove_slide(x)

# Remove multiple slides at once
x <- read_pptx()
x <- add_slide(x, "Title and Content")
x <- add_slide(
  x,
  layout = "Two Content",
  `Title 1` = "A title",
  dt = "Jan. 26, 2025",
  `body[2]` = "Body 2",
  left = "Left side",
  `6` = "Footer"
)
x <- add_slide(
  x,
  layout = "Two Content",
  `Title 1` = "A title",
  dt = "Jan. 26, 2025",
  `body[2]` = "Body 2",
  left = "Left side",
  `6` = "Footer"
)
x <- add_slide(x, "Title and Content")
x <- remove_slide(x, index = c(2, 4))
pptx_file <- print(x, target = tempfile(fileext = ".pptx"))
pptx_file
#> [1] "/tmp/RtmpFsafPq/file17984f58a040.pptx"
```
