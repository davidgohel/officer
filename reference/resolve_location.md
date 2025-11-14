# Resolve short form location

Convert short form location format, a numeric or a string (e.g.
`"body [1]"`), into its corresponding location object. Under the hood,
we parse the short form location and call the corresponding function
from the `ph_location_*` family. Note that short forms may not cover all
function from the `ph_location_*` and offer less customization.

## Usage

``` r
resolve_location(x)
```

## Short forms

The following location short forms are implemented. The corresponding
call of the function from the `ph_location_*` family is displayed on the
right.

|                |                                                             |                                                                                                    |
|----------------|-------------------------------------------------------------|----------------------------------------------------------------------------------------------------|
| **Short form** | **Description**                                             | **Location function**                                                                              |
| `"left"`       | Keyword string                                              | [`ph_location_left()`](https://davidgohel.github.io/officer/reference/ph_location_left.md)         |
| `"right"`      | Keyword string                                              | [`ph_location_right()`](https://davidgohel.github.io/officer/reference/ph_location_right.md)       |
| `"fullsize"`   | Keyword string                                              | [`ph_location_fullsize()`](https://davidgohel.github.io/officer/reference/ph_location_fullsize.md) |
| `"body [1]"`   | String: type + index in brackets (`1` if omitted)           | `ph_location_type("body", 1)`                                                                      |
| `"my_label"`   | Any string not matching a keyword or type                   | `ph_location_label("my_label")`                                                                    |
| `2`            | Length 1 integer                                            | `ph_location_id(2)`                                                                                |
| `c(0,0,4,5)`   | Length 4 numeric, optionally named, `c(top=0, left=0, ...)` | `ph_location(0, 0, 4, 5)`                                                                          |

## Examples

``` r
resolve_location("left")
#> $label
#> NULL
#> 
#> attr(,"class")
#> [1] "location_left" "location_str" 
resolve_location("right")
#> $label
#> NULL
#> 
#> attr(,"class")
#> [1] "location_right" "location_str"  
resolve_location("fullsize")
#> $label
#> [1] ""
#> 
#> attr(,"class")
#> [1] "location_fullsize" "location_str"     
resolve_location("body")
#> $type
#> [1] "body"
#> 
#> $type_idx
#> [1] 1
#> 
#> $position_right
#> [1] TRUE
#> 
#> $position_top
#> [1] TRUE
#> 
#> $id
#> NULL
#> 
#> $label
#> NULL
#> 
#> attr(,"class")
#> [1] "location_type" "location_str" 
resolve_location("body [1]")
#> $type
#> [1] "body"
#> 
#> $type_idx
#> [1] 1
#> 
#> $position_right
#> [1] TRUE
#> 
#> $position_top
#> [1] TRUE
#> 
#> $id
#> NULL
#> 
#> $label
#> NULL
#> 
#> attr(,"class")
#> [1] "location_type" "location_str" 
resolve_location("<some label>")
#> $ph_label
#> [1] "<some label>"
#> 
#> $label
#> NULL
#> 
#> attr(,"class")
#> [1] "location_label" "location_str"  
resolve_location(2)
#> $type
#> NULL
#> 
#> $type_idx
#> NULL
#> 
#> $position_right
#> NULL
#> 
#> $position_right
#> NULL
#> 
#> $position_top
#> NULL
#> 
#> $id
#> NULL
#> 
#> $ph_id
#> [1] 2
#> 
#> $label
#> NULL
#> 
#> attr(,"class")
#> [1] "location_id"  "location_num"
resolve_location(c(0, 0, 4, 5))
#> $left
#> [1] 0
#> 
#> $top
#> [1] 0
#> 
#> $width
#> [1] 4
#> 
#> $height
#> [1] 5
#> 
#> $ph_label
#> [1] ""
#> 
#> $ph
#> [1] NA
#> 
#> $bg
#> NULL
#> 
#> $rotation
#> NULL
#> 
#> $ln
#> NULL
#> 
#> $geom
#> NULL
#> 
#> $fld_type
#> [1] NA
#> 
#> $fld_id
#> [1] NA
#> 
#> attr(,"class")
#> [1] "location_manual" "location_str"   
```
