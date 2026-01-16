# Location of a named placeholder

The function will use the label of a placeholder to find the
corresponding location.

## Usage

``` r
ph_location_label(ph_label, newlabel = NULL, ...)
```

## Arguments

- ph_label:

  placeholder label of the used layout. It can be read in PowerPoint or
  with function
  [`layout_properties()`](https://davidgohel.github.io/officer/dev/reference/layout_properties.md)
  in column `ph_label`.

- newlabel:

  a label to associate with the placeholder.

- ...:

  unused arguments

## Details

The location of the bounding box associated to a placeholder within a
slide is specified with the left top coordinate, the width and the
height. These are defined in inches:

- left:

  left coordinate of the bounding box

- top:

  top coordinate of the bounding box

- width:

  width of the bounding box

- height:

  height of the bounding box

In addition to these attributes, a label can be associated with the
shape. Shapes, text boxes, images and other objects will be identified
with that label in the *Selection Pane* of PowerPoint. This label can
then be reused by other functions such as `ph_location_label()`. It can
be set with argument `newlabel`.

## See also

Other functions for placeholder location:
[`ph_location()`](https://davidgohel.github.io/officer/dev/reference/ph_location.md),
[`ph_location_fullsize()`](https://davidgohel.github.io/officer/dev/reference/ph_location_fullsize.md),
[`ph_location_id()`](https://davidgohel.github.io/officer/dev/reference/ph_location_id.md),
[`ph_location_left()`](https://davidgohel.github.io/officer/dev/reference/ph_location_left.md),
[`ph_location_right()`](https://davidgohel.github.io/officer/dev/reference/ph_location_right.md),
[`ph_location_template()`](https://davidgohel.github.io/officer/dev/reference/ph_location_template.md),
[`ph_location_type()`](https://davidgohel.github.io/officer/dev/reference/ph_location_type.md)

## Examples

``` r
library(officer)

# ph_location_label demo ----

doc <- read_pptx()
doc <- add_slide(doc, layout = "Title and Content")

# all ph_label can be read here
layout_properties(doc, layout = "Title and Content")
#>    master_name              name   type type_idx id                   ph_label
#> 1 Office Theme Title and Content  title        1  2                    Title 1
#> 2 Office Theme Title and Content   body        1  3      Content Placeholder 2
#> 3 Office Theme Title and Content     dt        1  4         Date Placeholder 3
#> 4 Office Theme Title and Content    ftr        1  5       Footer Placeholder 4
#> 5 Office Theme Title and Content sldNum        1  6 Slide Number Placeholder 5
#>                                            ph     offx      offy       cx
#> 1                        <p:ph type="title"/> 0.500000 0.3003478 9.000000
#> 2                             <p:ph idx="1"/> 0.500000 1.7500000 9.000000
#> 3        <p:ph type="dt" sz="half" idx="10"/> 0.500000 6.9513889 2.333333
#> 4    <p:ph type="ftr" sz="quarter" idx="11"/> 3.416667 6.9513889 3.166667
#> 5 <p:ph type="sldNum" sz="quarter" idx="12"/> 7.166667 6.9513889 2.333333
#>          cy rotation                                 fld_id          fld_type
#> 1 1.2500000       NA                                   <NA>              <NA>
#> 2 4.9496533       NA                                   <NA>              <NA>
#> 3 0.3993056       NA {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 4 0.3993056       NA                                   <NA>              <NA>
#> 5 0.3993056       NA {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum

doc <- ph_with(
  doc,
  head(iris),
  location = ph_location_label(ph_label = "Content Placeholder 2")
)
doc <- ph_with(
  doc,
  format(Sys.Date()),
  location = ph_location_label(ph_label = "Date Placeholder 3")
)
doc <- ph_with(
  doc,
  "This is a title",
  location = ph_location_label(ph_label = "Title 1")
)

print(doc, target = tempfile(fileext = ".pptx"))
```
