# Tabulation mark properties object

create a tabulation mark properties setting object for Word or RTF.
Results can be used as arguments of
[`fp_tabs()`](https://davidgohel.github.io/officer/reference/fp_tabs.md).

Once tabulation marks settings are defined, tabulation marks can be
added with
[`run_tab()`](https://davidgohel.github.io/officer/reference/run_tab.md)
inside a call to
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md) or
with `\t` within 'flextable' content.

## Usage

``` r
fp_tab(pos, style = "decimal")
```

## Arguments

- pos:

  Specifies the position of the tab stop (in inches).

- style:

  style of the tab. Possible values are: "decimal", "left", "right" or
  "center".

## See also

Other functions for defining formatting properties:
[`fp_border()`](https://davidgohel.github.io/officer/reference/fp_border.md),
[`fp_cell()`](https://davidgohel.github.io/officer/reference/fp_cell.md),
[`fp_par()`](https://davidgohel.github.io/officer/reference/fp_par.md),
[`fp_tabs()`](https://davidgohel.github.io/officer/reference/fp_tabs.md),
[`fp_text()`](https://davidgohel.github.io/officer/reference/fp_text.md)

## Examples

``` r
fp_tab(pos = 0.4, style = "decimal")
#> $pos
#> [1] 0.4
#> 
#> $style
#> [1] "decimal"
#> 
#> attr(,"class")
#> [1] "fp_tab"
fp_tab(pos = 1, style = "right")
#> $pos
#> [1] 1
#> 
#> $style
#> [1] "right"
#> 
#> attr(,"class")
#> [1] "fp_tab"
```
