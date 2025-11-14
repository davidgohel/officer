# Create a 'PowerPoint' document object

Read and import a pptx file as an R object representing the document.

The function is called `read_pptx` because it allows you to initialize
an object of class `rpptx` from an existing PowerPoint file. Content
will be added to the existing presentation. By default, an empty
document is used.

## Usage

``` r
read_pptx(path = NULL)
```

## Arguments

- path:

  path to the pptx file to use as base document. `potx` file are
  supported.

## master layouts and slide layouts

`read_pptx()` uses a PowerPoint file as the initial document. This is
the original PowerPoint document where all slide layouts, placeholders
for shapes and styles come from. Major points to be aware of are:

- Slide layouts are relative to a master layout. A document can contain
  one or more master layouts; a master layout can contain one or more
  slide layouts.

- A slide layout inherits design properties from its master layout but
  some properties can be overwritten.

- Designs and formatting properties of layouts and shapes (placeholders
  in a layout) are defined within the initial document. There is no R
  function to modify these values - they must be defined in the initial
  document.

## See also

[`print.rpptx()`](https://davidgohel.github.io/officer/reference/print.rpptx.md),
[`add_slide()`](https://davidgohel.github.io/officer/reference/add_slide.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/reference/plot_layout_properties.md),
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md)

## Examples

``` r
read_pptx()
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
```
