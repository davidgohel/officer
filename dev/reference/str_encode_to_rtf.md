# Encode UTF8 string to RTF

Convert strings to RTF valid codes.

## Usage

``` r
str_encode_to_rtf(z)
```

## Arguments

- z:

  character vector to be converted

## Value

character vector of results encoded to RTF

## See also

Other functions for officer extensions:
[`fortify_location()`](https://davidgohel.github.io/officer/dev/reference/fortify_location.md),
[`get_reference_value()`](https://davidgohel.github.io/officer/dev/reference/get_reference_value.md),
[`opts_current_table()`](https://davidgohel.github.io/officer/dev/reference/opts_current_table.md),
[`shape_properties_tags()`](https://davidgohel.github.io/officer/dev/reference/shape_properties_tags.md),
[`to_html()`](https://davidgohel.github.io/officer/dev/reference/to_html.md),
[`to_pml()`](https://davidgohel.github.io/officer/dev/reference/to_pml.md),
[`to_rtf()`](https://davidgohel.github.io/officer/dev/reference/to_rtf.md),
[`to_wml()`](https://davidgohel.github.io/officer/dev/reference/to_wml.md),
[`wml_link_images()`](https://davidgohel.github.io/officer/dev/reference/wml_link_images.md)

## Examples

``` r
str_encode_to_rtf("Hello World")
#> [1] "Hello World"
```
