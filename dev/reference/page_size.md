# Page size object

The function creates a representation of the dimensions of a page. The
dimensions are defined by length, width and orientation. If the
orientation is in landscape mode then the length becomes the width and
the width becomes the length.

## Usage

``` r
page_size(
  width = 11906/1440,
  height = 16838/1440,
  orient = "portrait",
  unit = "in"
)
```

## Arguments

- width, height:

  page width, page height, default to A4 format If NULL the value will
  be ignored and Word will use the default value.

- orient:

  page orientation, either 'landscape', either 'portrait'.

- unit:

  unit for width and height, one of "in", "cm", "mm".

## See also

Other functions for section definition:
[`page_mar()`](https://davidgohel.github.io/officer/dev/reference/page_mar.md),
[`prop_section()`](https://davidgohel.github.io/officer/dev/reference/prop_section.md),
[`section_columns()`](https://davidgohel.github.io/officer/dev/reference/section_columns.md)

## Examples

``` r
page_size(orient = "landscape")
#> $width
#> [1] 11.69306
#> 
#> $height
#> [1] 8.268056
#> 
#> $orient
#> [1] "landscape"
#> 
#> attr(,"class")
#> [1] "page_size"
```
