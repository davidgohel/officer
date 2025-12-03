# Section for 'Word'

Create a representation of a section.

A section affects preceding paragraphs or tables; i.e. a section starts
at the end of the previous section (or the beginning of the document if
no preceding section exists), and stops where the section is declared.

When a new landscape section is needed, it is recommended to add a
block_section with `type = "continuous"`, to add the content to be
appened in the new section and finally to add a block_section with
`page_size = page_size(orient = "landscape")`.

## Usage

``` r
block_section(property)
```

## Arguments

- property:

  section properties defined with function
  [prop_section](https://davidgohel.github.io/officer/dev/reference/prop_section.md)

## See also

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/dev/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/dev/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/dev/reference/block_pour_docx.md),
[`block_table()`](https://davidgohel.github.io/officer/dev/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/dev/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/dev/reference/unordered_list.md)

## Examples

``` r
ps <- prop_section(
  page_size = page_size(orient = "landscape"),
  page_margins = page_mar(top = 2),
  type = "continuous"
)
block_section(ps)
#> ----- end of section: 
```
