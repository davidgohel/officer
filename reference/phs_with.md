# Fill multiple placeholders using key value syntax

A sibling of
[ph_with](https://davidgohel.github.io/officer/reference/ph_with.md)
that fills mutiple placeholders at once. Placeholder locations are
specfied using the short form syntax. The location and corresponding
object are passed as key value pairs
(`phs_with("short form location" = object)`). Under the hood,
[ph_with](https://davidgohel.github.io/officer/reference/ph_with.md) is
called for each pair. Note that `phs_with` does not cover all options
from the `ph_location_*` family and is also less customization. It is a
covenience wrapper for the most common use cases. The implemented short
forms are listed in section `"Short forms"`.

## Usage

``` r
phs_with(x, ..., .dots = NULL, .slide_idx = NULL)
```

## Arguments

- x:

  A `rpptx` object.

- ...:

  Key-value pairs of the form `"short form location" = object`. If the
  short form is an integer or a string with blanks, you must wrap it in
  quotes or backticks.

- .dots:

  List of key-value pairs `"short form location" = object`. Alternative
  to `...`.

- .slide_idx:

  Numeric indexes of slides to process. `NULL` (default) processes the
  current slide only. Use keyword `all` for all slides.

## Short forms

The following short forms are implemented and can be used as the
parameter in the function call. The corresponding function from the
`ph_location_*` family (called under the hood) is displayed on the
right.

|                |                                                   |                                                                                                    |
|----------------|---------------------------------------------------|----------------------------------------------------------------------------------------------------|
| **Short form** | **Description**                                   | **Location function**                                                                              |
| `"left"`       | Keyword string                                    | [`ph_location_left()`](https://davidgohel.github.io/officer/reference/ph_location_left.md)         |
| `"right"`      | Keyword string                                    | [`ph_location_right()`](https://davidgohel.github.io/officer/reference/ph_location_right.md)       |
| `"fullsize"`   | Keyword string                                    | [`ph_location_fullsize()`](https://davidgohel.github.io/officer/reference/ph_location_fullsize.md) |
| `"body [1]"`   | String: type + index in brackets (`1` if omitted) | `ph_location_type("body", 1)`                                                                      |
| `"my_label"`   | Any string not matching a keyword or type         | `ph_location_label("my_label")`                                                                    |
| `1`            | Length 1 integer                                  | `ph_location_id(1)`                                                                                |

## See also

[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md),
[`add_slide()`](https://davidgohel.github.io/officer/reference/add_slide.md)

## Examples

``` r
library(officer)

# use key-value format to fill phs
x <- read_pptx()
x <- add_slide(x, "Two Content")
x <- phs_with(
  x,
  `Title 1` = "A title", # ph label
  dt = Sys.Date(), # ph type
  `body[2]` = "Body 2", # ph type + type index
  left = "Left side", # ph keyword
  `6` = "Footer" # ph index
)

# reuse ph content via the .dots arg
x <- read_pptx()
my_ph_list <- list(`6` = "Footer", dt = Sys.Date())
x <- add_slide(x, "Two Content")
x <- phs_with(
  x,
  `Title 1` = "Title A",
  `body[2]` = "Body A",
  .dots = my_ph_list
)
x <- add_slide(x, "Two Content")
x <- phs_with(
  x,
  `Title 1` = "Title B",
  `body[2]` = "Body B",
  .dots = my_ph_list
)

# use the .slide_idx arg to select which slide(s) to process
x <- read_pptx()
x <- add_slide(x, "Two Content")
x <- add_slide(x, "Two Content")
x <- phs_with(x, `6` = "Footer", dt = Sys.Date(), .slide_idx = 1:2)

# run to open temp pptx file locally
# \dontrun{
# print(x, preview = TRUE)
# }
```
