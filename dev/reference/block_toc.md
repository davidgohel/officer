# Table of content for 'Word'

Create a representation of a table of content for Word documents.

## Usage

``` r
block_toc(level = 3, style = NULL, seq_id = NULL, separator = ";")
```

## Arguments

- level:

  max title level of the table

- style:

  optional. If not NULL, its value is used as style in the document that
  will be used to build entries of the TOC.

- seq_id:

  optional. If not NULL, its value is used as sequence identifier in the
  document that will be used to build entries of the TOC. See also
  [`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md)
  to specify a sequence identifier.

- separator:

  unused, no effect

## See also

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/dev/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/dev/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/dev/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/dev/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/dev/reference/block_table.md),
[`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/dev/reference/unordered_list.md)

## Examples

``` r
block_toc(level = 2)
#> TOC - max level: 2
block_toc(style = "Table Caption")
#> TOC for style: Table Caption
```
