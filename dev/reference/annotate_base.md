# Placeholder parameters annotation

generates a slide from each layout in the base document to identify the
placeholder indexes, types, names, master names and layout names.

This is to be used when need to know what parameters should be used with
`ph_location*` calls. The parameters are printed in their corresponding
shapes.

Note that if there are duplicated `ph_label`, you should not use
[`ph_location_label()`](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md).
Hint: You can dedupe labels using
[`layout_dedupe_ph_labels()`](https://davidgohel.github.io/officer/dev/reference/layout_dedupe_ph_labels.md).

## Usage

``` r
annotate_base(path = NULL, output_file = "annotated_layout.pptx")
```

## Arguments

- path:

  path to the pptx file to use as base document or NULL to use the
  officer default

- output_file:

  filename to store the annotated powerpoint file or NULL to suppress
  generation

## Value

rpptx object of the annotated PowerPoint file

## See also

Other functions for reading presentation information:
[`color_scheme()`](https://davidgohel.github.io/officer/dev/reference/color_scheme.md),
[`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md),
[`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md),
[`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md),
[`length.rpptx()`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md),
[`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)

## Examples

``` r
# To generate an anotation of the default base document with officer:
annotate_base(output_file = tempfile(fileext = ".pptx"))
#> pptx document with 7 slides
#> Available layouts and their associated master(s):
#>              layout       master
#> 1       Title Slide Office Theme
#> 2 Title and Content Office Theme
#> 3    Section Header Office Theme
#> 4       Two Content Office Theme
#> 5        Comparison Office Theme
#> 6        Title Only Office Theme
#> 7             Blank Office Theme

# To generate an annotation of the base document 'mydoc.pptx' and place the
# annotated output in 'mydoc_annotate.pptx'
# annotate_base(path = 'mydoc.pptx', output_file='mydoc_annotate.pptx')
```
