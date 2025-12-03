# Tabs properties object

create a set of tabulation mark properties object for Word or RTF.
Results can be used as arguments `tabs` of
[`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)
and will only have effects in Word or RTF outputs.

Once a set of tabulation marks settings is defined, tabulation marks can
be added with
[`run_tab()`](https://davidgohel.github.io/officer/dev/reference/run_tab.md)
inside a call to
[`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md)
or with `\t` within 'flextable' content.

## Usage

``` r
fp_tabs(...)
```

## Arguments

- ...:

  [fp_tab](https://davidgohel.github.io/officer/dev/reference/fp_tab.md)
  objects

## See also

Other functions for defining formatting properties:
[`fp_border()`](https://davidgohel.github.io/officer/dev/reference/fp_border.md),
[`fp_cell()`](https://davidgohel.github.io/officer/dev/reference/fp_cell.md),
[`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md),
[`fp_tab()`](https://davidgohel.github.io/officer/dev/reference/fp_tab.md),
[`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)

## Examples

``` r
z <- fp_tabs(
  fp_tab(pos = 0.4, style = "decimal"),
  fp_tab(pos = 1, style = "decimal")
)
fpar(
  run_tab(), ftext("88."),
  run_tab(), ftext("987.45"),
  fp_p = fp_par(
    tabs = z
  )
)
#> $chunks
#> $chunks[[1]]
#> list()
#> attr(,"class")
#> [1] "run_tab" "run"    
#> 
#> $chunks[[2]]
#> text: 88.
#> format:
#> NULL
#> 
#> $chunks[[3]]
#> list()
#> attr(,"class")
#> [1] "run_tab" "run"    
#> 
#> $chunks[[4]]
#> text: 987.45
#> format:
#> NULL
#> 
#> 
#> $fp_p
#>                     values
#> text.align            left
#> padding.top              0
#> padding.bottom           0
#> padding.left             0
#> padding.right            0
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
#> 
#> $fp_t
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        NA     NA   NA         NA     NA    NA      NA       NA          NA
#>   fontname_eastasia fontname.hansi vertical_align
#> 1                NA             NA       baseline
#> 
#> attr(,"class")
#> [1] "fpar"  "block"
```
