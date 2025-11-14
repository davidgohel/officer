# Text formatting properties

Create an `fp_text` object that describes text formatting properties.

Function `fp_text_lite()` is generating properties with only entries for
the parameters users provided. The undefined properties will inherit
from the default settings.

## Usage

``` r
fp_text(
  color = "black",
  font.size = 10,
  bold = FALSE,
  italic = FALSE,
  underlined = FALSE,
  strike = FALSE,
  font.family = "Arial",
  cs.family = NULL,
  eastasia.family = NULL,
  hansi.family = NULL,
  vertical.align = "baseline",
  shading.color = "transparent"
)

fp_text_lite(
  color = NA,
  font.size = NA,
  font.family = NA,
  cs.family = NA,
  eastasia.family = NA,
  hansi.family = NA,
  bold = NA,
  italic = NA,
  underlined = NA,
  strike = NA,
  vertical.align = "baseline",
  shading.color = NA
)

# S3 method for class 'fp_text'
format(x, type = "wml", ...)

# S3 method for class 'fp_text'
print(x, ...)

# S3 method for class 'fp_text'
update(
  object,
  color,
  font.size,
  bold,
  italic,
  underlined,
  strike,
  font.family,
  cs.family,
  eastasia.family,
  hansi.family,
  vertical.align,
  shading.color,
  ...
)
```

## Arguments

- color:

  font color - a single character value specifying a valid color (e.g.
  "#000000" or "black").

- font.size:

  font size (in point) - 0 or positive integer value.

- bold:

  is bold

- italic:

  is italic

- underlined:

  is underlined

- strike:

  is strikethrough

- font.family:

  single character value. Specifies the font to be used to format
  characters in the Unicode range (U+0000-U+007F).

- cs.family:

  optional font to be used to format characters in a complex script
  Unicode range. For example, Arabic text might be displayed using the
  "Arial Unicode MS" font.

- eastasia.family:

  optional font to be used to format characters in an East Asian Unicode
  range. For example, Japanese text might be displayed using the "MS
  Mincho" font.

- hansi.family:

  optional. Specifies the font to be used to format characters in a
  Unicode range which does not fall into one of the other categories.

- vertical.align:

  single character value specifying font vertical alignments. Expected
  value is one of the following : default `'baseline'` or `'subscript'`
  or `'superscript'`

- shading.color:

  shading color - a single character value specifying a valid color
  (e.g. "#000000" or "black").

- x:

  `fp_text` object

- type:

  output type - one of 'wml', 'pml', 'html', 'rtf'.

- ...:

  further arguments - not used

- object:

  `fp_text` object to modify

- format:

  format type, wml for MS word, pml for MS PowerPoint and html.

## Value

a `fp_text` object

## See also

[`ftext()`](https://davidgohel.github.io/officer/reference/ftext.md),
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md)

Other functions for defining formatting properties:
[`fp_border()`](https://davidgohel.github.io/officer/reference/fp_border.md),
[`fp_cell()`](https://davidgohel.github.io/officer/reference/fp_cell.md),
[`fp_par()`](https://davidgohel.github.io/officer/reference/fp_par.md),
[`fp_tab()`](https://davidgohel.github.io/officer/reference/fp_tab.md),
[`fp_tabs()`](https://davidgohel.github.io/officer/reference/fp_tabs.md)

## Examples

``` r
fp_text()
#>   font.size italic  bold underlined strike color     shading fontname
#> 1        10  FALSE FALSE      FALSE  FALSE black transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
fp_text(color = "red")
#>   font.size italic  bold underlined strike color     shading fontname
#> 1        10  FALSE FALSE      FALSE  FALSE   red transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
fp_text(bold = TRUE, shading.color = "yellow")
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        10  FALSE TRUE      FALSE  FALSE black  yellow    Arial       Arial
#>   fontname_eastasia fontname.hansi vertical_align
#> 1             Arial          Arial       baseline
print(fp_text(color = "red", font.size = 12))
#>   font.size italic  bold underlined strike color     shading fontname
#> 1        12  FALSE FALSE      FALSE  FALSE   red transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
```
