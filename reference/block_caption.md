# Caption block

Create a representation of a caption that can be used for cross
reference.

## Usage

``` r
block_caption(label, style = NULL, autonum = NULL)
```

## Arguments

- label:

  a scalar character representing label to display

- style:

  paragraph style name

- autonum:

  an object generated with function
  [run_autonum](https://davidgohel.github.io/officer/reference/run_autonum.md)

## See also

Other block functions for reporting:
[`block_gg()`](https://davidgohel.github.io/officer/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/reference/unordered_list.md)

## Examples

``` r
library(officer)

run_num <- run_autonum(seq_id = "tab", pre_label = "tab. ",
  bkm = "mtcars_table")
caption <- block_caption("mtcars table",
  style = "Normal",
  autonum = run_num
)

doc_1 <- read_docx()
doc_1 <- body_add(doc_1, "A title", style = "heading 1")
doc_1 <- body_add(doc_1, "Hello world!", style = "Normal")
doc_1 <- body_add(doc_1, caption)
doc_1 <- body_add(doc_1, mtcars, style = "table_template")

print(doc_1, target = tempfile(fileext = ".docx"))
```
