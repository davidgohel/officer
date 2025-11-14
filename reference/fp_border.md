# Border properties object

create a border properties object.

## Usage

``` r
fp_border(color = "black", style = "solid", width = 1)

# S3 method for class 'fp_border'
update(object, color, style, width, ...)
```

## Arguments

- color:

  border color - single character value (e.g. "#000000" or "black")

- style:

  border style - single character value : See Details for supported
  border styles.

- width:

  border width - an integer value : 0\>= value

- object:

  fp_border object

- ...:

  further arguments - not used

## Details

For Word output the following border styles are supported:

- "none" or "nil" - No Border

- "solid" or "single" - Single Line Border

- "thick" - Single Line Border

- "double" - Double Line Border

- "dotted" - Dotted Line Border

- "dashed" - Dashed Line Border

- "dotDash" - Dot Dash Line Border

- "dotDotDash" - Dot Dot Dash Line Border

- "triple" - Triple Line Border

- "thinThickSmallGap" - Thin, Thick Line Border

- "thickThinSmallGap" - Thick, Thin Line Border

- "thinThickThinSmallGap" - Thin, Thick, Thin Line Border

- "thinThickMediumGap" - Thin, Thick Line Border

- "thickThinMediumGap" - Thick, Thin Line Border

- "thinThickThinMediumGap" - Thin, Thick, Thin Line Border

- "thinThickLargeGap" - Thin, Thick Line Border

- "thickThinLargeGap" - Thick, Thin Line Border

- "thinThickThinLargeGap" - Thin, Thick, Thin Line Border

- "wave" - Wavy Line Border

- "doubleWave" - Double Wave Line Border

- "dashSmallGap" - Dashed Line Border

- "dashDotStroked" - Dash Dot Strokes Line Border

- "threeDEmboss" or "ridge" - 3D Embossed Line Border

- "threeDEngrave" or "groove" - 3D Engraved Line Border

- "outset" - Outset Line Border

- "inset" - Inset Line Border

For HTML output only a limited amount of border styles are supported:

- "none" or "nil" - No Border

- "solid" or "single" - Single Line Border

- "double" - Double Line Border

- "dotted" - Dotted Line Border

- "dashed" - Dashed Line Border

- "threeDEmboss" or "ridge" - 3D Embossed Line Border

- "threeDEngrave" or "groove" - 3D Engraved Line Border

- "outset" - Outset Line Border

- "inset" - Inset Line Border

Non-supported Word border styles will default to "solid".

## See also

Other functions for defining formatting properties:
[`fp_cell()`](https://davidgohel.github.io/officer/reference/fp_cell.md),
[`fp_par()`](https://davidgohel.github.io/officer/reference/fp_par.md),
[`fp_tab()`](https://davidgohel.github.io/officer/reference/fp_tab.md),
[`fp_tabs()`](https://davidgohel.github.io/officer/reference/fp_tabs.md),
[`fp_text()`](https://davidgohel.github.io/officer/reference/fp_text.md)

## Examples

``` r
fp_border()
#> line: color: black, width: 1, style: solid
fp_border(color = "orange", style = "solid", width = 1)
#> line: color: orange, width: 1, style: solid
fp_border(color = "gray", style = "dotted", width = 1)
#> line: color: gray, width: 1, style: dotted

# modify object ------
border <- fp_border()
update(border, style = "dotted", width = 3)
#> line: color: black, width: 3, style: dotted
```
