# Slide content in a data.frame

Get content and positions of current slide into a data.frame. Data for
any tables, images, or paragraphs are imported into the resulting
data.frame.

## Usage

``` r
slide_summary(x, index = NULL)
```

## Arguments

- x:

  an rpptx object

- index:

  slide index

## Note

The column `id` of the result is not to be used by users. This is a
technical string id whose value will be used by office when the document
will be rendered. This is not related to argument `index` required by
functions
[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md).

## See also

Other functions for reading presentation information:
[`annotate_base()`](https://davidgohel.github.io/officer/dev/reference/annotate_base.md),
[`color_scheme()`](https://davidgohel.github.io/officer/dev/reference/color_scheme.md),
[`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md),
[`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md),
[`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md),
[`length.rpptx()`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md)

## Examples

``` r
my_pres <- read_pptx()
my_pres <- add_slide(my_pres, "Title and Content")
my_pres <- ph_with(my_pres, format(Sys.Date()),
  location = ph_location_type(type="dt"))
my_pres <- add_slide(my_pres, "Title and Content")
my_pres <- ph_with(my_pres, iris[1:2,],
  location = ph_location_type(type="body"))
slide_summary(my_pres)
#>   type                                   id              ph_label offx offy cx
#> 1 body 4e5d4136-cdc7-439f-86ea-7e7ef93e7af6 Content Placeholder 2   NA   NA NA
#>   cy rotation fld_id fld_type
#> 1 NA       NA   <NA>     <NA>
#>                                                                                                                              text
#> 1 {5C22544A-7EE6-4342-B048-85BDC9FD1C3A}Sepal.LengthSepal.WidthPetal.LengthPetal.WidthSpecies5.13.51.40.2setosa4.93.01.40.2setosa
slide_summary(my_pres, index = 1)
#>   type                                   id           ph_label offx     offy
#> 1   dt b9f1fbc8-c853-487c-85df-e4baa654b08c Date Placeholder 3  0.5 6.951389
#>         cx        cy rotation fld_id fld_type       text
#> 1 2.333333 0.3993056       NA   <NA>     <NA> 2026-01-16
```
