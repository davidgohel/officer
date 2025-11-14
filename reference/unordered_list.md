# Unordered list

unordered list of text for PowerPoint presentations. Each text is
associated with a hierarchy level.

## Usage

``` r
unordered_list(str_list = character(0), level_list = integer(0), style = NULL)
```

## Arguments

- str_list:

  list of strings to be included in the object

- level_list:

  list of levels for hierarchy structure. Use 0 for 'no bullet', 1 for
  level 1, 2 for level 2 and so on.

- style:

  text style, a `fp_text` object list or a single `fp_text` objects. Use
  `fp_text(font.size = 0, ...)` to inherit from default sizes of the
  presentation.

## See also

[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md)

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/reference/plot_instr.md)

## Examples

``` r
unordered_list(
level_list = c(1, 2, 2, 3, 3, 1),
str_list = c("Level1", "Level2", "Level2", "Level3", "Level3", "Level1"),
style = fp_text(color = "red", font.size = 0) )
#>      str lvl
#> 1 Level1   1
#> 2 Level2   2
#> 3 Level2   2
#> 4 Level3   3
#> 5 Level3   3
#> 6 Level1   1
unordered_list(
level_list = c(1, 2, 1),
str_list = c("Level1", "Level2", "Level1"),
style = list(
  fp_text(color = "red", font.size = 0),
  fp_text(color = "pink", font.size = 0),
  fp_text(color = "orange", font.size = 0)
  ))
#>      str lvl
#> 1 Level1   1
#> 2 Level2   2
#> 3 Level1   1
```
