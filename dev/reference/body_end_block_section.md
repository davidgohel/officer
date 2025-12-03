# Add any section

Add a section to the document. You can define any section with a
[block_section](https://davidgohel.github.io/officer/dev/reference/block_section.md)
object. All other `body_end_section_*` are specialized, this one is
highly flexible but it's up to the user to define the section
properties.

## Usage

``` r
body_end_block_section(x, value)
```

## Arguments

- x:

  an rdocx object

- value:

  a
  [block_section](https://davidgohel.github.io/officer/dev/reference/block_section.md)
  object

## Illustrations

![](figures/body_end_block_section_doc_1.png)

## See also

Other functions for Word sections:
[`body_end_section_columns()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_columns.md),
[`body_end_section_columns_landscape()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_columns_landscape.md),
[`body_end_section_continuous()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_continuous.md),
[`body_end_section_landscape()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_landscape.md),
[`body_end_section_portrait()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_portrait.md),
[`body_set_default_section()`](https://davidgohel.github.io/officer/dev/reference/body_set_default_section.md)

## Examples

``` r
library(officer)
str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
str1 <- rep(str1, 20)
str1 <- paste(str1, collapse = " ")

ps <- prop_section(
  page_size = page_size(orient = "landscape"),
  page_margins = page_mar(top = 2),
  type = "continuous"
)

doc_1 <- read_docx()
doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")

doc_1 <- body_end_block_section(doc_1, block_section(ps))

doc_1 <- body_add_par(doc_1, value = str1, style = "centered")

print(doc_1, target = tempfile(fileext = ".docx"))
```
