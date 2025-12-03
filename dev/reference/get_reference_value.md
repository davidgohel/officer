# Get the document being used as a template

Get filename of the document being used as a template in an R Markdown
document rendered as HTML, PowerPoint presentation or Word document. It
requires packages rmarkdown \>= 1.10.14 and knitr.

## Usage

``` r
get_reference_value(format = NULL)
```

## Arguments

- format:

  document format, one of 'pptx', 'docx' or 'html'

## Value

a name file

## See also

Other functions for officer extensions:
[`fortify_location()`](https://davidgohel.github.io/officer/dev/reference/fortify_location.md),
[`opts_current_table()`](https://davidgohel.github.io/officer/dev/reference/opts_current_table.md),
[`shape_properties_tags()`](https://davidgohel.github.io/officer/dev/reference/shape_properties_tags.md),
[`str_encode_to_rtf()`](https://davidgohel.github.io/officer/dev/reference/str_encode_to_rtf.md),
[`to_html()`](https://davidgohel.github.io/officer/dev/reference/to_html.md),
[`to_pml()`](https://davidgohel.github.io/officer/dev/reference/to_pml.md),
[`to_rtf()`](https://davidgohel.github.io/officer/dev/reference/to_rtf.md),
[`to_wml()`](https://davidgohel.github.io/officer/dev/reference/to_wml.md),
[`wml_link_images()`](https://davidgohel.github.io/officer/dev/reference/wml_link_images.md)
