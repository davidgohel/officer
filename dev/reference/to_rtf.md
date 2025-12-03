# Convert officer objects to RTF

Convert an object made with package officer to RTF.

## Usage

``` r
to_rtf(x, ...)
```

## Arguments

- x:

  object to convert to RTF. Supported objects are:

  - [`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md)

  - [`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md)

  - [`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md)

  - [`run_columnbreak()`](https://davidgohel.github.io/officer/dev/reference/run_columnbreak.md)

  - [`run_linebreak()`](https://davidgohel.github.io/officer/dev/reference/run_linebreak.md)

  - [`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md)

  - [`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md)

  - [`run_pagebreak()`](https://davidgohel.github.io/officer/dev/reference/run_pagebreak.md)

  - [`hyperlink_ftext()`](https://davidgohel.github.io/officer/dev/reference/hyperlink_ftext.md)

  - [`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md)

  - [`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md)

- ...:

  Arguments to be passed to methods

## Value

a string containing the RTF code

## See also

Other functions for officer extensions:
[`fortify_location()`](https://davidgohel.github.io/officer/dev/reference/fortify_location.md),
[`get_reference_value()`](https://davidgohel.github.io/officer/dev/reference/get_reference_value.md),
[`opts_current_table()`](https://davidgohel.github.io/officer/dev/reference/opts_current_table.md),
[`shape_properties_tags()`](https://davidgohel.github.io/officer/dev/reference/shape_properties_tags.md),
[`str_encode_to_rtf()`](https://davidgohel.github.io/officer/dev/reference/str_encode_to_rtf.md),
[`to_html()`](https://davidgohel.github.io/officer/dev/reference/to_html.md),
[`to_pml()`](https://davidgohel.github.io/officer/dev/reference/to_pml.md),
[`to_wml()`](https://davidgohel.github.io/officer/dev/reference/to_wml.md),
[`wml_link_images()`](https://davidgohel.github.io/officer/dev/reference/wml_link_images.md)
