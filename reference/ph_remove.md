# Remove a shape

Remove a shape in a slide.

## Usage

``` r
ph_remove(x, type = "body", id = 1, ph_label = NULL, id_chr = NULL)
```

## Arguments

- x:

  an rpptx object

- type:

  placeholder type

- id:

  placeholder index (integer) for a duplicated type. This is to be used
  when a placeholder type is not unique in the layout of the current
  slide, e.g. two placeholders with type 'body'. To add onto the first,
  use `id = 1` and `id = 2` for the second one. Values can be read from
  [`slide_summary()`](https://davidgohel.github.io/officer/reference/slide_summary.md).

- ph_label:

  label associated to the placeholder. Use column `ph_label` of result
  returned by
  [`slide_summary()`](https://davidgohel.github.io/officer/reference/slide_summary.md).
  If used, `type` and `id` are ignored.

- id_chr:

  deprecated.

## See also

[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md)

Other functions for placeholders manipulation:
[`ph_hyperlink()`](https://davidgohel.github.io/officer/reference/ph_hyperlink.md),
[`ph_slidelink()`](https://davidgohel.github.io/officer/reference/ph_slidelink.md)

## Examples

``` r
fileout <- tempfile(fileext = ".pptx")
dummy_fun <- function(doc) {
  doc <- add_slide(doc,
    layout = "Two Content",
    master = "Office Theme"
  )
  doc <- ph_with(
    x = doc, value = "Un titre",
    location = ph_location_type(type = "title")
  )
  doc <- ph_with(
    x = doc, value = "Un corps 1",
    location = ph_location_type(type = "body", id = 1)
  )
  doc <- ph_with(
    x = doc, value = "Un corps 2",
    location = ph_location_type(type = "body", id = 2)
  )
  doc
}
doc <- read_pptx()
for (i in 1:3) {
  doc <- dummy_fun(doc)
}
#> Warning: ! The `id` argument in `ph_location_type()` is deprecated as of officer 0.6.7.
#> ℹ Please use the `type_idx` argument instead.
#> ✖ Caution: new index logic in `type_idx` (see docs).
#> Warning: ! The `id` argument in `ph_location_type()` is deprecated as of officer 0.6.7.
#> ℹ Please use the `type_idx` argument instead.
#> ✖ Caution: new index logic in `type_idx` (see docs).
#> Warning: ! The `id` argument in `ph_location_type()` is deprecated as of officer 0.6.7.
#> ℹ Please use the `type_idx` argument instead.
#> ✖ Caution: new index logic in `type_idx` (see docs).
#> Warning: ! The `id` argument in `ph_location_type()` is deprecated as of officer 0.6.7.
#> ℹ Please use the `type_idx` argument instead.
#> ✖ Caution: new index logic in `type_idx` (see docs).
#> Warning: ! The `id` argument in `ph_location_type()` is deprecated as of officer 0.6.7.
#> ℹ Please use the `type_idx` argument instead.
#> ✖ Caution: new index logic in `type_idx` (see docs).
#> Warning: ! The `id` argument in `ph_location_type()` is deprecated as of officer 0.6.7.
#> ℹ Please use the `type_idx` argument instead.
#> ✖ Caution: new index logic in `type_idx` (see docs).

doc <- on_slide(doc, index = 1)
doc <- ph_remove(x = doc, type = "title")

doc <- on_slide(doc, index = 2)
doc <- ph_remove(x = doc, type = "body", id = 2)

doc <- on_slide(doc, index = 3)
doc <- ph_remove(x = doc, type = "body", id = 1)

print(doc, target = fileout)
```
