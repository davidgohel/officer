# Add a landscape multi columns section

A landscape section with multiple columns is added to the document.

## Usage

``` r
body_end_section_columns_landscape(
  x,
  widths = c(2.5, 2.5),
  space = 0.25,
  sep = FALSE,
  w = 16838/1440,
  h = 11906/1440
)
```

## Arguments

- x:

  an rdocx object

- widths:

  columns widths in inches. If 3 values, 3 columns will be produced.

- space:

  space in inches between columns.

- sep:

  if TRUE a line is separating columns.

- w, h:

  page width, page height (in inches)

## Section breaks occupy a line

Ending a section is materialized in the document by an empty paragraph
mark holding the section properties; this is how the OOXML format
represents a section break, there is no other way, and Word does the
same when a section break is inserted manually. As any paragraph mark,
it occupies one line whose height depends on the default paragraph
style. The same applies to
[`run_columnbreak()`](https://davidgohel.github.io/officer/dev/reference/run_columnbreak.md)
when it has to be placed after a table: a table cannot host the break,
so it must live in a paragraph that also takes one line. If that
residual line matters for your layout, define a dedicated paragraph
style with a very small font size (e.g. 1pt) and no spacing in your Word
template, and use it for the paragraph hosting the break:
`fpar(run_columnbreak(), fp_p = fp_par(word_style = "MyTinyStyle"))`.

## See also

Other functions for Word sections:
[`body_end_block_section()`](https://davidgohel.github.io/officer/dev/reference/body_end_block_section.md),
[`body_end_section_columns()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_columns.md),
[`body_end_section_continuous()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_continuous.md),
[`body_end_section_landscape()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_landscape.md),
[`body_end_section_portrait()`](https://davidgohel.github.io/officer/dev/reference/body_end_section_portrait.md),
[`body_set_default_section()`](https://davidgohel.github.io/officer/dev/reference/body_set_default_section.md)

## Examples

``` r
str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
str1 <- rep(str1, 5)
str1 <- paste(str1, collapse = " ")

doc_1 <- read_docx()
doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
doc_1 <- body_end_section_columns_landscape(doc_1, widths = c(6, 2))
doc_1 <- body_add_par(doc_1, value = str1, style = "Normal")
print(doc_1, target = tempfile(fileext = ".docx"))
```
