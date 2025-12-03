# Line end properties

Create a `sp_lineend` object that describes line end properties.

## Usage

``` r
sp_lineend(type = "none", width = "med", length = "med")

# S3 method for class 'sp_lineend'
print(x, ...)

# S3 method for class 'sp_lineend'
update(object, type, width, length, ...)
```

## Arguments

- type:

  single character value specifying the line end type. Expected value is
  one of the following : default `'none'` or `'triangle'` or `'stealth'`
  or `'diamond'` or `'oval'` or `'arrow'`

- width:

  single character value specifying the line end width Expected value is
  one of the following : default `'sm'` or `'med'` or `'lg'`

- length:

  single character value specifying the line end length Expected value
  is one of the following : default `'sm'` or `'med'` or `'lg'`

- x, object:

  `sp_lineend` object

- ...:

  further arguments - not used

## Value

a `sp_lineend` object

## See also

[sp_line](https://davidgohel.github.io/officer/dev/reference/sp_line.md)

Other functions for defining shape properties:
[`sp_line()`](https://davidgohel.github.io/officer/dev/reference/sp_line.md)

## Examples

``` r
sp_lineend()
#>   type width length
#> 1 none   med    med
sp_lineend(type = "triangle")
#>       type width length
#> 1 triangle   med    med
sp_lineend(type = "arrow", width = "lg", length = "lg")
#>    type width length
#> 1 arrow    lg     lg
print(sp_lineend (type="triangle", width = "lg"))
#>       type width length
#> 1 triangle    lg    med
obj <- sp_lineend (type="triangle", width = "lg")
update( obj, type = "arrow" )
#>    type width length
#> 1 arrow    lg    med
```
