# Line properties

Create a `sp_line` object that describes line properties.

## Usage

``` r
sp_line(
  color = "transparent",
  lwd = 1,
  lty = "solid",
  linecmpd = "sng",
  lineend = "rnd",
  linejoin = "round",
  headend = sp_lineend(type = "none"),
  tailend = sp_lineend(type = "none")
)

# S3 method for class 'sp_line'
print(x, ...)

# S3 method for class 'sp_line'
update(
  object,
  color,
  lwd,
  lty,
  linecmpd,
  lineend,
  linejoin,
  headend,
  tailend,
  ...
)
```

## Arguments

- color:

  line color - a single character value specifying a valid color (e.g.
  "#000000" or "black").

- lwd:

  line width (in point) - 0 or positive integer value.

- lty:

  single character value specifying the line type. Expected value is one
  of the following : default `'solid'` or `'dot'` or `'dash'` or
  `'lgDash'` or `'dashDot'` or `'lgDashDot'` or `'lgDashDotDot'` or
  `'sysDash'` or `'sysDot'` or `'sysDashDot'` or `'sysDashDotDot'`.

- linecmpd:

  single character value specifying the compound line type. Expected
  value is one of the following : default `'sng'` or `'dbl'` or `'tri'`
  or `'thinThick'` or `'thickThin'`

- lineend:

  single character value specifying the line end style Expected value is
  one of the following : default `'rnd'` or `'sq'` or `'flat'`

- linejoin:

  single character value specifying the line join style Expected value
  is one of the following : default `'round'` or `'bevel'` or `'miter'`

- headend:

  a `sp_lineend` object specifying line head end style

- tailend:

  a `sp_lineend` object specifying line tail end style

- x, object:

  `sp_line` object

- ...:

  further arguments - not used

## Value

a `sp_line` object

## See also

[sp_lineend](https://davidgohel.github.io/officer/dev/reference/sp_lineend.md)

Other functions for defining shape properties:
[`sp_lineend()`](https://davidgohel.github.io/officer/dev/reference/sp_lineend.md)

## Examples

``` r
sp_line()
#>         color lwd   lty linecmpd lineend linejoin headend.type headend.width
#> 1 transparent   1 solid      sng     rnd    round         none           med
#>   headend.length tailend.type tailend.width tailend.length
#> 1            med         none           med            med
sp_line(color = "red", lwd = 2)
#>   color lwd   lty linecmpd lineend linejoin headend.type headend.width
#> 1   red   2 solid      sng     rnd    round         none           med
#>   headend.length tailend.type tailend.width tailend.length
#> 1            med         none           med            med
sp_line(lty = "dot", linecmpd = "dbl")
#>         color lwd lty linecmpd lineend linejoin headend.type headend.width
#> 1 transparent   1 dot      dbl     rnd    round         none           med
#>   headend.length tailend.type tailend.width tailend.length
#> 1            med         none           med            med
print( sp_line (color="red", lwd = 2) )
#>   color lwd   lty linecmpd lineend linejoin headend.type headend.width
#> 1   red   2 solid      sng     rnd    round         none           med
#>   headend.length tailend.type tailend.width tailend.length
#> 1            med         none           med            med
obj <- sp_line (color="red", lwd = 2)
update( obj, linecmpd = "dbl" )
#>   color lwd   lty linecmpd lineend linejoin headend.type headend.width
#> 1   red   2 solid      dbl     rnd    round         none           med
#>   headend.length tailend.type tailend.width tailend.length
#> 1            med         none           med            med
```
