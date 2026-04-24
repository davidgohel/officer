# List of items for Word and PowerPoint

Create a bullet or numbered list from
[`list_item()`](https://davidgohel.github.io/officer/reference/list_item.md)
elements. Supports rich text via
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md) and
multi-level nesting. Works in both Word and PowerPoint documents.

## Usage

``` r
block_list_items(..., list_type = "bullet")
```

## Arguments

- ...:

  [`list_item()`](https://davidgohel.github.io/officer/reference/list_item.md)
  objects

- list_type:

  `"bullet"` for an unordered list or `"decimal"` for a numbered list

## See also

[`list_item()`](https://davidgohel.github.io/officer/reference/list_item.md),
[`body_add()`](https://davidgohel.github.io/officer/reference/body_add.md),
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
[`list_item()`](https://davidgohel.github.io/officer/reference/list_item.md),
[`plot_instr()`](https://davidgohel.github.io/officer/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/reference/unordered_list.md)

## Examples

``` r
items <- block_list_items(
  list_item(fpar(
    ftext("Item 1", fp_text(color = "red"))
  ), level = 1),
  list_item(fpar("Sub-item"), level = 2),
  list_item(fpar("Item 2"), level = 1),
  list_type = "bullet"
)
items
#> Bullet list (3 items):
#>   - Item 1
#>     - Sub-item
#>   - Item 2
```
