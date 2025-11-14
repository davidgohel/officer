# Eval a location on the current slide

Eval a shape location against the current slide. This function is to be
used to add custom openxml code. A list is returned, it contains
informations width, height, left and top positions and other
informations necessary to add a content on a slide.

## Usage

``` r
fortify_location(x, doc, ...)
```

## Arguments

- x:

  a location for a placeholder.

- doc:

  an rpptx object

- ...:

  unused arguments

## See also

[`ph_location()`](https://davidgohel.github.io/officer/reference/ph_location.md),
[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md)

Other functions for officer extensions:
[`get_reference_value()`](https://davidgohel.github.io/officer/reference/get_reference_value.md),
[`opts_current_table()`](https://davidgohel.github.io/officer/reference/opts_current_table.md),
[`shape_properties_tags()`](https://davidgohel.github.io/officer/reference/shape_properties_tags.md),
[`str_encode_to_rtf()`](https://davidgohel.github.io/officer/reference/str_encode_to_rtf.md),
[`to_html()`](https://davidgohel.github.io/officer/reference/to_html.md),
[`to_pml()`](https://davidgohel.github.io/officer/reference/to_pml.md),
[`to_rtf()`](https://davidgohel.github.io/officer/reference/to_rtf.md),
[`to_wml()`](https://davidgohel.github.io/officer/reference/to_wml.md),
[`wml_link_images()`](https://davidgohel.github.io/officer/reference/wml_link_images.md)

## Examples

``` r
doc <- read_pptx()
doc <- add_slide(doc,
  layout = "Title and Content",
  master = "Office Theme"
)
fortify_location(ph_location_fullsize(), doc)
#> $width
#> [1] 10
#> 
#> $height
#> [1] 7.5
#> 
#> $left
#> [1] 0
#> 
#> $top
#> [1] 0
#> 
#> $ph_label
#> [1] ""
#> 
#> $ph
#> [1] NA
#> 
#> $type
#> [1] "body"
#> 
#> $rotation
#> [1] 0
#> 
#> $fld_id
#> [1] NA
#> 
#> $fld_type
#> [1] NA
#> 
```
