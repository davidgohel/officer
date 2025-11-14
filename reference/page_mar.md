# Page margins object

Define margins for each page of a section.

The function creates a representation of the dimensions of a page. The
dimensions are defined by length, width and orientation. If the
orientation is in landscape mode then the length becomes the width and
the width becomes the length.

## Usage

``` r
page_mar(
  bottom = 1417/1440,
  top = 1417/1440,
  right = 1417/1440,
  left = 1417/1440,
  header = 708/1440,
  footer = 708/1440,
  gutter = 0/1440
)
```

## Arguments

- bottom, top:

  distance (in inches) between the bottom/top of the text margin and the
  bottom/top of the page. The text is placed at the greater of the value
  of this attribute and the extent of the header/footer text. A negative
  value indicates that the content should be measured from the
  bottom/top of the page regardless of the footer/header, and so will
  overlap the footer/header. For example, `header=-0.5, bottom=1` means
  that the footer must start one inch from the bottom of the page and
  the main document text must start a half inch from the bottom of the
  page. In this case, the text and footer overlap since bottom is
  negative.

- left, right:

  distance (in inches) from the left/right edge of the page to the
  left/right edge of the text.

- header:

  distance (in inches) from the top edge of the page to the top edge of
  the header.

- footer:

  distance (in inches) from the bottom edge of the page to the bottom
  edge of the footer.

- gutter:

  page gutter (in inches).

## See also

Other functions for section definition:
[`page_size()`](https://davidgohel.github.io/officer/reference/page_size.md),
[`prop_section()`](https://davidgohel.github.io/officer/reference/prop_section.md),
[`section_columns()`](https://davidgohel.github.io/officer/reference/section_columns.md)

## Examples

``` r
page_mar()
#> $header
#> [1] 0.4916667
#> 
#> $bottom
#> [1] 0.9840278
#> 
#> $top
#> [1] 0.9840278
#> 
#> $right
#> [1] 0.9840278
#> 
#> $left
#> [1] 0.9840278
#> 
#> $footer
#> [1] 0.4916667
#> 
#> $gutter
#> [1] 0
#> 
#> attr(,"class")
#> [1] "page_mar"
```
