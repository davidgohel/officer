# Hyperlink a placeholder

Add hyperlink to a placeholder in the current slide.

## Usage

``` r
ph_hyperlink(x, type = "body", id = 1, id_chr = NULL, ph_label = NULL, href)
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

- id_chr:

  deprecated.

- ph_label:

  label associated to the placeholder. Use column `ph_label` of result
  returned by
  [`slide_summary()`](https://davidgohel.github.io/officer/reference/slide_summary.md).
  If used, `type` and `id` are ignored.

- href:

  hyperlink (do not forget http or https prefix)

## See also

[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md)

Other functions for placeholders manipulation:
[`ph_remove()`](https://davidgohel.github.io/officer/reference/ph_remove.md),
[`ph_slidelink()`](https://davidgohel.github.io/officer/reference/ph_slidelink.md)

## Examples

``` r
fileout <- tempfile(fileext = ".pptx")
loc_manual <- ph_location(bg = "red", newlabel = "mytitle")
doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(x = doc, "Un titre 1", location = loc_manual)
slide_summary(doc) # read column ph_label here
#>   type                                   id ph_label offx offy cx cy rotation
#> 1 body 6c77cf13-1a16-4bd2-acc7-4dab1c2c77a3  mytitle    1    1  4  3       NA
#>   fld_id fld_type       text
#> 1   <NA>     <NA> Un titre 1
doc <- ph_hyperlink(
  x = doc, ph_label = "mytitle",
  href = "https://cran.r-project.org"
)

print(doc, target = fileout)
```
