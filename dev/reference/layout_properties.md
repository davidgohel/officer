# Slide layout properties

Detailed information about the placeholders on the slide layouts (label,
position, etc.). See *Value* section below for more info.

## Usage

``` r
layout_properties(x, layout = NULL, master = NULL)
```

## Arguments

- x:

  an `rpptx` object

- layout:

  slide layout name. If `NULL`, returns all layouts.

- master:

  master layout name where `layout` is located. If `NULL`, returns all
  masters.

## Value

Returns a data frame with one row per placeholder and the following
columns:

- `master_name`: Name of master (a `.pptx` file may have more than one)

- `name`: Name of layout

- `type`: Placeholder type

- `type_idx`: Running index for phs of the same type. Ordering by ph
  position (top -\> bottom, left -\> right)

- `id`: A unique placeholder id (assigned by PowerPoint automatically,
  starts at 2, potentially non-consecutive)

- `ph_label`: Placeholder label (can be set by the user in PowerPoint)

- `ph`: Placholder XML fragment (usually not needed)

- `offx`,`offy`: placeholder's distance from left and top edge (in inch)

- `cx`,`cy`: width and height of placeholder (in inch)

- `rotation`: rotation in degrees

- `fld_id` is generally stored as a hexadecimal or GUID value

- `fld_type`: a unique identifier for a particular field

## See also

Other functions for reading presentation information:
[`annotate_base()`](https://davidgohel.github.io/officer/dev/reference/annotate_base.md),
[`color_scheme()`](https://davidgohel.github.io/officer/dev/reference/color_scheme.md),
[`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md),
[`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md),
[`length.rpptx()`](https://davidgohel.github.io/officer/dev/reference/length.rpptx.md),
[`plot_layout_properties()`](https://davidgohel.github.io/officer/dev/reference/plot_layout_properties.md),
[`slide_size()`](https://davidgohel.github.io/officer/dev/reference/slide_size.md),
[`slide_summary()`](https://davidgohel.github.io/officer/dev/reference/slide_summary.md)

## Examples

``` r
x <- read_pptx()
layout_properties(x = x, layout = "Title Slide", master = "Office Theme")
#>    master_name        name     type type_idx id                   ph_label
#> 1 Office Theme Title Slide ctrTitle        1  2                    Title 1
#> 2 Office Theme Title Slide subTitle        1  3                 Subtitle 2
#> 3 Office Theme Title Slide       dt        1  4         Date Placeholder 3
#> 4 Office Theme Title Slide      ftr        1  5       Footer Placeholder 4
#> 5 Office Theme Title Slide   sldNum        1  6 Slide Number Placeholder 5
#>                                            ph     offx     offy       cx
#> 1                     <p:ph type="ctrTitle"/> 0.750000 2.329861 8.500000
#> 2             <p:ph type="subTitle" idx="1"/> 1.500000 4.250000 7.000000
#> 3        <p:ph type="dt" sz="half" idx="10"/> 0.500000 6.951389 2.333333
#> 4    <p:ph type="ftr" sz="quarter" idx="11"/> 3.416667 6.951389 3.166667
#> 5 <p:ph type="sldNum" sz="quarter" idx="12"/> 7.166667 6.951389 2.333333
#>          cy rotation                                 fld_id          fld_type
#> 1 1.6076389       NA                                   <NA>              <NA>
#> 2 1.9166667       NA                                   <NA>              <NA>
#> 3 0.3993056       NA {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 4 0.3993056       NA                                   <NA>              <NA>
#> 5 0.3993056       NA {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
layout_properties(x = x, master = "Office Theme")
#>     master_name              name     type type_idx id
#> 1  Office Theme       Title Slide ctrTitle        1  2
#> 2  Office Theme       Title Slide subTitle        1  3
#> 3  Office Theme       Title Slide       dt        1  4
#> 4  Office Theme       Title Slide      ftr        1  5
#> 5  Office Theme       Title Slide   sldNum        1  6
#> 6  Office Theme Title and Content    title        1  2
#> 7  Office Theme Title and Content     body        1  3
#> 8  Office Theme Title and Content       dt        1  4
#> 9  Office Theme Title and Content      ftr        1  5
#> 10 Office Theme Title and Content   sldNum        1  6
#> 11 Office Theme    Section Header     body        1  3
#> 12 Office Theme    Section Header    title        1  2
#> 13 Office Theme    Section Header       dt        1  4
#> 14 Office Theme    Section Header      ftr        1  5
#> 15 Office Theme    Section Header   sldNum        1  6
#> 16 Office Theme       Two Content    title        1  2
#> 17 Office Theme       Two Content     body        1  3
#> 18 Office Theme       Two Content     body        2  4
#> 19 Office Theme       Two Content       dt        1  5
#> 20 Office Theme       Two Content      ftr        1  6
#> 21 Office Theme       Two Content   sldNum        1  7
#> 22 Office Theme        Comparison    title        1  2
#> 23 Office Theme        Comparison     body        1  3
#> 24 Office Theme        Comparison     body        2  5
#> 25 Office Theme        Comparison     body        3  4
#> 26 Office Theme        Comparison     body        4  6
#> 27 Office Theme        Comparison       dt        1  7
#> 28 Office Theme        Comparison      ftr        1  8
#> 29 Office Theme        Comparison   sldNum        1  9
#> 30 Office Theme        Title Only    title        1  2
#> 31 Office Theme        Title Only       dt        1  3
#> 32 Office Theme        Title Only      ftr        1  4
#> 33 Office Theme        Title Only   sldNum        1  5
#> 34 Office Theme             Blank       dt        1  2
#> 35 Office Theme             Blank      ftr        1  3
#> 36 Office Theme             Blank   sldNum        1  4
#>                      ph_label                                          ph
#> 1                     Title 1                     <p:ph type="ctrTitle"/>
#> 2                  Subtitle 2             <p:ph type="subTitle" idx="1"/>
#> 3          Date Placeholder 3        <p:ph type="dt" sz="half" idx="10"/>
#> 4        Footer Placeholder 4    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 5  Slide Number Placeholder 5 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 6                     Title 1                        <p:ph type="title"/>
#> 7       Content Placeholder 2                             <p:ph idx="1"/>
#> 8          Date Placeholder 3        <p:ph type="dt" sz="half" idx="10"/>
#> 9        Footer Placeholder 4    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 10 Slide Number Placeholder 5 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 11         Text Placeholder 2                 <p:ph type="body" idx="1"/>
#> 12                    Title 1                        <p:ph type="title"/>
#> 13         Date Placeholder 3        <p:ph type="dt" sz="half" idx="10"/>
#> 14       Footer Placeholder 4    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 15 Slide Number Placeholder 5 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 16                    Title 1                        <p:ph type="title"/>
#> 17      Content Placeholder 2                   <p:ph sz="half" idx="1"/>
#> 18      Content Placeholder 3                   <p:ph sz="half" idx="2"/>
#> 19         Date Placeholder 4        <p:ph type="dt" sz="half" idx="10"/>
#> 20       Footer Placeholder 5    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 21 Slide Number Placeholder 6 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 22                    Title 1                        <p:ph type="title"/>
#> 23         Text Placeholder 2                 <p:ph type="body" idx="1"/>
#> 24         Text Placeholder 4    <p:ph type="body" sz="quarter" idx="3"/>
#> 25      Content Placeholder 3                   <p:ph sz="half" idx="2"/>
#> 26      Content Placeholder 5                <p:ph sz="quarter" idx="4"/>
#> 27         Date Placeholder 6        <p:ph type="dt" sz="half" idx="10"/>
#> 28       Footer Placeholder 7    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 29 Slide Number Placeholder 8 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 30                    Title 1                        <p:ph type="title"/>
#> 31         Date Placeholder 2        <p:ph type="dt" sz="half" idx="10"/>
#> 32       Footer Placeholder 3    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 33 Slide Number Placeholder 4 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 34         Date Placeholder 1        <p:ph type="dt" sz="half" idx="10"/>
#> 35       Footer Placeholder 2    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 36 Slide Number Placeholder 3 <p:ph type="sldNum" sz="quarter" idx="12"/>
#>         offx      offy       cx        cy rotation
#> 1  0.7500000 2.3298611 8.500000 1.6076389       NA
#> 2  1.5000000 4.2500000 7.000000 1.9166667       NA
#> 3  0.5000000 6.9513889 2.333333 0.3993056       NA
#> 4  3.4166667 6.9513889 3.166667 0.3993056       NA
#> 5  7.1666667 6.9513889 2.333333 0.3993056       NA
#> 6  0.5000000 0.3003478 9.000000 1.2500000       NA
#> 7  0.5000000 1.7500000 9.000000 4.9496533       NA
#> 8  0.5000000 6.9513889 2.333333 0.3993056       NA
#> 9  3.4166667 6.9513889 3.166667 0.3993056       NA
#> 10 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 11 0.7899311 3.1788200 8.500000 1.6406245       NA
#> 12 0.7899311 4.8194444 8.500000 1.4895833       NA
#> 13 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 14 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 15 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 16 0.5000000 0.3003478 9.000000 1.2500000       NA
#> 17 0.5000000 1.7500000 4.416667 4.9496533       NA
#> 18 5.0833333 1.7500000 4.416667 4.9496533       NA
#> 19 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 20 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 21 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 22 0.5000000 0.3003478 9.000000 1.2500000       NA
#> 23 0.5000000 1.6788200 4.418403 0.6996522       NA
#> 24 5.0798611 1.6788200 4.420139 0.6996522       NA
#> 25 0.5000000 2.3784722 4.418403 4.3211811       NA
#> 26 5.0798611 2.3784722 4.420139 4.3211811       NA
#> 27 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 28 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 29 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 30 0.5000000 0.3003478 9.000000 1.2500000       NA
#> 31 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 32 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 33 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 34 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 35 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 36 7.1666667 6.9513889 2.333333 0.3993056       NA
#>                                    fld_id          fld_type
#> 1                                    <NA>              <NA>
#> 2                                    <NA>              <NA>
#> 3  {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 4                                    <NA>              <NA>
#> 5  {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 6                                    <NA>              <NA>
#> 7                                    <NA>              <NA>
#> 8  {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 9                                    <NA>              <NA>
#> 10 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 11                                   <NA>              <NA>
#> 12                                   <NA>              <NA>
#> 13 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 14                                   <NA>              <NA>
#> 15 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 16                                   <NA>              <NA>
#> 17                                   <NA>              <NA>
#> 18                                   <NA>              <NA>
#> 19 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 20                                   <NA>              <NA>
#> 21 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 22                                   <NA>              <NA>
#> 23                                   <NA>              <NA>
#> 24                                   <NA>              <NA>
#> 25                                   <NA>              <NA>
#> 26                                   <NA>              <NA>
#> 27 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 28                                   <NA>              <NA>
#> 29 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 30                                   <NA>              <NA>
#> 31 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 32                                   <NA>              <NA>
#> 33 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 34 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 35                                   <NA>              <NA>
#> 36 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
layout_properties(x = x, layout = "Two Content")
#>    master_name        name   type type_idx id                   ph_label
#> 1 Office Theme Two Content  title        1  2                    Title 1
#> 2 Office Theme Two Content   body        1  3      Content Placeholder 2
#> 3 Office Theme Two Content   body        2  4      Content Placeholder 3
#> 4 Office Theme Two Content     dt        1  5         Date Placeholder 4
#> 5 Office Theme Two Content    ftr        1  6       Footer Placeholder 5
#> 6 Office Theme Two Content sldNum        1  7 Slide Number Placeholder 6
#>                                            ph     offx      offy       cx
#> 1                        <p:ph type="title"/> 0.500000 0.3003478 9.000000
#> 2                   <p:ph sz="half" idx="1"/> 0.500000 1.7500000 4.416667
#> 3                   <p:ph sz="half" idx="2"/> 5.083333 1.7500000 4.416667
#> 4        <p:ph type="dt" sz="half" idx="10"/> 0.500000 6.9513889 2.333333
#> 5    <p:ph type="ftr" sz="quarter" idx="11"/> 3.416667 6.9513889 3.166667
#> 6 <p:ph type="sldNum" sz="quarter" idx="12"/> 7.166667 6.9513889 2.333333
#>          cy rotation                                 fld_id          fld_type
#> 1 1.2500000       NA                                   <NA>              <NA>
#> 2 4.9496533       NA                                   <NA>              <NA>
#> 3 4.9496533       NA                                   <NA>              <NA>
#> 4 0.3993056       NA {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 5 0.3993056       NA                                   <NA>              <NA>
#> 6 0.3993056       NA {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
layout_properties(x = x)
#>     master_name              name     type type_idx id
#> 1  Office Theme       Title Slide ctrTitle        1  2
#> 2  Office Theme       Title Slide subTitle        1  3
#> 3  Office Theme       Title Slide       dt        1  4
#> 4  Office Theme       Title Slide      ftr        1  5
#> 5  Office Theme       Title Slide   sldNum        1  6
#> 6  Office Theme Title and Content    title        1  2
#> 7  Office Theme Title and Content     body        1  3
#> 8  Office Theme Title and Content       dt        1  4
#> 9  Office Theme Title and Content      ftr        1  5
#> 10 Office Theme Title and Content   sldNum        1  6
#> 11 Office Theme    Section Header     body        1  3
#> 12 Office Theme    Section Header    title        1  2
#> 13 Office Theme    Section Header       dt        1  4
#> 14 Office Theme    Section Header      ftr        1  5
#> 15 Office Theme    Section Header   sldNum        1  6
#> 16 Office Theme       Two Content    title        1  2
#> 17 Office Theme       Two Content     body        1  3
#> 18 Office Theme       Two Content     body        2  4
#> 19 Office Theme       Two Content       dt        1  5
#> 20 Office Theme       Two Content      ftr        1  6
#> 21 Office Theme       Two Content   sldNum        1  7
#> 22 Office Theme        Comparison    title        1  2
#> 23 Office Theme        Comparison     body        1  3
#> 24 Office Theme        Comparison     body        2  5
#> 25 Office Theme        Comparison     body        3  4
#> 26 Office Theme        Comparison     body        4  6
#> 27 Office Theme        Comparison       dt        1  7
#> 28 Office Theme        Comparison      ftr        1  8
#> 29 Office Theme        Comparison   sldNum        1  9
#> 30 Office Theme        Title Only    title        1  2
#> 31 Office Theme        Title Only       dt        1  3
#> 32 Office Theme        Title Only      ftr        1  4
#> 33 Office Theme        Title Only   sldNum        1  5
#> 34 Office Theme             Blank       dt        1  2
#> 35 Office Theme             Blank      ftr        1  3
#> 36 Office Theme             Blank   sldNum        1  4
#>                      ph_label                                          ph
#> 1                     Title 1                     <p:ph type="ctrTitle"/>
#> 2                  Subtitle 2             <p:ph type="subTitle" idx="1"/>
#> 3          Date Placeholder 3        <p:ph type="dt" sz="half" idx="10"/>
#> 4        Footer Placeholder 4    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 5  Slide Number Placeholder 5 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 6                     Title 1                        <p:ph type="title"/>
#> 7       Content Placeholder 2                             <p:ph idx="1"/>
#> 8          Date Placeholder 3        <p:ph type="dt" sz="half" idx="10"/>
#> 9        Footer Placeholder 4    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 10 Slide Number Placeholder 5 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 11         Text Placeholder 2                 <p:ph type="body" idx="1"/>
#> 12                    Title 1                        <p:ph type="title"/>
#> 13         Date Placeholder 3        <p:ph type="dt" sz="half" idx="10"/>
#> 14       Footer Placeholder 4    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 15 Slide Number Placeholder 5 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 16                    Title 1                        <p:ph type="title"/>
#> 17      Content Placeholder 2                   <p:ph sz="half" idx="1"/>
#> 18      Content Placeholder 3                   <p:ph sz="half" idx="2"/>
#> 19         Date Placeholder 4        <p:ph type="dt" sz="half" idx="10"/>
#> 20       Footer Placeholder 5    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 21 Slide Number Placeholder 6 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 22                    Title 1                        <p:ph type="title"/>
#> 23         Text Placeholder 2                 <p:ph type="body" idx="1"/>
#> 24         Text Placeholder 4    <p:ph type="body" sz="quarter" idx="3"/>
#> 25      Content Placeholder 3                   <p:ph sz="half" idx="2"/>
#> 26      Content Placeholder 5                <p:ph sz="quarter" idx="4"/>
#> 27         Date Placeholder 6        <p:ph type="dt" sz="half" idx="10"/>
#> 28       Footer Placeholder 7    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 29 Slide Number Placeholder 8 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 30                    Title 1                        <p:ph type="title"/>
#> 31         Date Placeholder 2        <p:ph type="dt" sz="half" idx="10"/>
#> 32       Footer Placeholder 3    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 33 Slide Number Placeholder 4 <p:ph type="sldNum" sz="quarter" idx="12"/>
#> 34         Date Placeholder 1        <p:ph type="dt" sz="half" idx="10"/>
#> 35       Footer Placeholder 2    <p:ph type="ftr" sz="quarter" idx="11"/>
#> 36 Slide Number Placeholder 3 <p:ph type="sldNum" sz="quarter" idx="12"/>
#>         offx      offy       cx        cy rotation
#> 1  0.7500000 2.3298611 8.500000 1.6076389       NA
#> 2  1.5000000 4.2500000 7.000000 1.9166667       NA
#> 3  0.5000000 6.9513889 2.333333 0.3993056       NA
#> 4  3.4166667 6.9513889 3.166667 0.3993056       NA
#> 5  7.1666667 6.9513889 2.333333 0.3993056       NA
#> 6  0.5000000 0.3003478 9.000000 1.2500000       NA
#> 7  0.5000000 1.7500000 9.000000 4.9496533       NA
#> 8  0.5000000 6.9513889 2.333333 0.3993056       NA
#> 9  3.4166667 6.9513889 3.166667 0.3993056       NA
#> 10 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 11 0.7899311 3.1788200 8.500000 1.6406245       NA
#> 12 0.7899311 4.8194444 8.500000 1.4895833       NA
#> 13 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 14 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 15 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 16 0.5000000 0.3003478 9.000000 1.2500000       NA
#> 17 0.5000000 1.7500000 4.416667 4.9496533       NA
#> 18 5.0833333 1.7500000 4.416667 4.9496533       NA
#> 19 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 20 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 21 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 22 0.5000000 0.3003478 9.000000 1.2500000       NA
#> 23 0.5000000 1.6788200 4.418403 0.6996522       NA
#> 24 5.0798611 1.6788200 4.420139 0.6996522       NA
#> 25 0.5000000 2.3784722 4.418403 4.3211811       NA
#> 26 5.0798611 2.3784722 4.420139 4.3211811       NA
#> 27 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 28 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 29 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 30 0.5000000 0.3003478 9.000000 1.2500000       NA
#> 31 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 32 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 33 7.1666667 6.9513889 2.333333 0.3993056       NA
#> 34 0.5000000 6.9513889 2.333333 0.3993056       NA
#> 35 3.4166667 6.9513889 3.166667 0.3993056       NA
#> 36 7.1666667 6.9513889 2.333333 0.3993056       NA
#>                                    fld_id          fld_type
#> 1                                    <NA>              <NA>
#> 2                                    <NA>              <NA>
#> 3  {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 4                                    <NA>              <NA>
#> 5  {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 6                                    <NA>              <NA>
#> 7                                    <NA>              <NA>
#> 8  {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 9                                    <NA>              <NA>
#> 10 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 11                                   <NA>              <NA>
#> 12                                   <NA>              <NA>
#> 13 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 14                                   <NA>              <NA>
#> 15 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 16                                   <NA>              <NA>
#> 17                                   <NA>              <NA>
#> 18                                   <NA>              <NA>
#> 19 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 20                                   <NA>              <NA>
#> 21 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 22                                   <NA>              <NA>
#> 23                                   <NA>              <NA>
#> 24                                   <NA>              <NA>
#> 25                                   <NA>              <NA>
#> 26                                   <NA>              <NA>
#> 27 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 28                                   <NA>              <NA>
#> 29 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 30                                   <NA>              <NA>
#> 31 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 32                                   <NA>              <NA>
#> 33 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
#> 34 {E6744CE3-0875-4B69-89C0-6F72D8139561} datetimeFigureOut
#> 35                                   <NA>              <NA>
#> 36 {8DADB20D-508E-4C6D-A9E4-257D5607B0F6}          slidenum
```
