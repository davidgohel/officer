# Slide link to a placeholder

Add slide link to a placeholder in the current slide.

## Usage

``` r
ph_slidelink(
  x,
  type = "body",
  id = 1,
  id_chr = NULL,
  ph_label = NULL,
  slide_index
)
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
  [`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md).

- id_chr:

  deprecated.

- ph_label:

  label associated to the placeholder. Use column `ph_label` of result
  returned by
  [`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md).
  If used, `type` and `id` are ignored.

- slide_index:

  slide index to reach

## See also

[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)

Other functions for placeholders manipulation:
[`ph_hyperlink()`](https://davidgohel.github.io/officer/dev/reference/ph_hyperlink.md),
[`ph_remove()`](https://davidgohel.github.io/officer/dev/reference/ph_remove.md)

## Examples

``` r
fileout <- tempfile(fileext = ".pptx")
loc_title <- ph_location_type(type = "title")
doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(x = doc, "Un titre 1", location = loc_title)
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(x = doc, "Un titre 2", location = loc_title)
doc <- on_slide(doc, 1)
slide_summary(doc) # read column ph_label here
#>    type                                   id ph_label offx      offy cx   cy
#> 1 title d29505c9-2443-4eda-970e-5270e1177faf  Title 1  0.5 0.3003478  9 1.25
#>   rotation fld_id fld_type       text
#> 1       NA   <NA>     <NA> Un titre 1
doc <- ph_slidelink(x = doc, ph_label = "Title 1", slide_index = 2)

print(doc, target = fileout)
```
