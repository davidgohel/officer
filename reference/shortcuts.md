# shortcuts for formatting properties

Shortcuts for
[`fp_text()`](https://davidgohel.github.io/officer/reference/fp_text.md),
[`fp_par()`](https://davidgohel.github.io/officer/reference/fp_par.md),
[`fp_cell()`](https://davidgohel.github.io/officer/reference/fp_cell.md)
and
[`fp_border()`](https://davidgohel.github.io/officer/reference/fp_border.md).

## Usage

``` r
shortcuts
```

## Examples

``` r
shortcuts$fp_bold()
#>   font.size italic bold underlined strike color     shading fontname
#> 1        10  FALSE TRUE      FALSE  FALSE black transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
shortcuts$fp_italic()
#>   font.size italic  bold underlined strike color     shading fontname
#> 1        10   TRUE FALSE      FALSE  FALSE black transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
shortcuts$b_null()
#> line: color: black, width: 0, style: solid
```
