# Default layout for new slides

Set or remove the default layout used when calling
[`add_slide()`](https://davidgohel.github.io/officer/dev/reference/add_slide.md).

## Usage

``` r
layout_default(x, layout = NULL, master = NULL, as_list = FALSE)
```

## Arguments

- x:

  An `rpptx` object.

- layout:

  Layout name. If `NULL` (default), removes the default layout.

- master:

  Name of master. Only required if layout name is not unique across
  masters.

- as_list:

  If `TRUE`, return a list with layout and master instead of the `rpptx`
  object.

## Value

The `rpptx` object.

## See also

[`add_slide()`](https://davidgohel.github.io/officer/dev/reference/add_slide.md)

## Examples

``` r
# set and remove the default layout
x <- read_pptx()
layout_default(x) # no defaults
#> pptx document with 0 slides
#> Available layouts and their associated master(s):
#>              layout       master
#> 1       Title Slide Office Theme
#> 2 Title and Content Office Theme
#> 3    Section Header Office Theme
#> 4       Two Content Office Theme
#> 5        Comparison Office Theme
#> 6        Title Only Office Theme
#> 7             Blank Office Theme
x <- layout_default(x, "Title and Content") # set default
layout_default(x)
#> pptx document with 0 slides
#> Available layouts and their associated master(s):
#>              layout       master
#> 1       Title Slide Office Theme
#> 2 Title and Content Office Theme
#> 3    Section Header Office Theme
#> 4       Two Content Office Theme
#> 5        Comparison Office Theme
#> 6        Title Only Office Theme
#> 7             Blank Office Theme
x <- add_slide(x) # new slide with default layout
x <- layout_default(x, NULL) # remove default
layout_default(x) # no defaults
#> pptx document with 1 slide
#> Available layouts and their associated master(s):
#>              layout       master
#> 1       Title Slide Office Theme
#> 2 Title and Content Office Theme
#> 3    Section Header Office Theme
#> 4       Two Content Office Theme
#> 5        Comparison Office Theme
#> 6        Title Only Office Theme
#> 7             Blank Office Theme

# use when repeatedly adding slides with same layout
x <- read_pptx()
x <- layout_default(x, "Title and Content")
x <- add_slide(x, title = "1. Slide", body = "Some content")
x <- add_slide(x, title = "2. Slide", body = "Some more content")
x <- add_slide(x, title = "3. Slide", body = "Even more content")
```
