# Paragraph styles for columns

The function defines the paragraph styles for columns.

## Usage

``` r
table_stylenames(stylenames = list())
```

## Arguments

- stylenames:

  a named character vector, names are column names, values are paragraph
  styles associated with each column. If a column is not specified,
  default value 'Normal' is used. Another form is as a named list, the
  list names are the styles and the contents are column names to be
  formatted with the corresponding style.

## See also

Other functions for table definition:
[`prop_table()`](https://davidgohel.github.io/officer/reference/prop_table.md),
[`table_colwidths()`](https://davidgohel.github.io/officer/reference/table_colwidths.md),
[`table_conditional_formatting()`](https://davidgohel.github.io/officer/reference/table_conditional_formatting.md),
[`table_layout()`](https://davidgohel.github.io/officer/reference/table_layout.md),
[`table_width()`](https://davidgohel.github.io/officer/reference/table_width.md)

## Examples

``` r
library(officer)

stylenames <- c(
  vs = "centered", am = "centered",
  gear = "centered", carb = "centered"
)

doc_1 <- read_docx()
doc_1 <- body_add_table(doc_1,
  value = mtcars, style = "table_template",
  stylenames = table_stylenames(stylenames = stylenames)
)

print(doc_1, target = tempfile(fileext = ".docx"))


stylenames <- list(
  "centered" = c("vs", "am", "gear", "carb")
)

doc_2 <- read_docx()
doc_2 <- body_add_table(doc_2,
  value = mtcars, style = "table_template",
  stylenames = table_stylenames(stylenames = stylenames)
)

print(doc_2, target = tempfile(fileext = ".docx"))
```
