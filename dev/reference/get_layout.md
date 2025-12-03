# Layout selection helper

Select a layout by name or index. The master name is inferred and only
required for disambiguation in case the layout name is not unique across
masters.

## Usage

``` r
get_layout(
  x,
  layout = NULL,
  master = NULL,
  layout_by_id = TRUE,
  get_first = FALSE
)
```

## Arguments

- x:

  An `rpptx` object.

- layout:

  Layout name or index. Index refers to the row index of the
  [`layout_summary()`](https://davidgohel.github.io/officer/dev/reference/layout_summary.md)
  output.

- master:

  Name of master. Only required if layout name is not unique across
  masters.

- layout_by_id:

  Allow layout index instead of name? (default is `TRUE`)

- get_first:

  If layout exists in multiple master, return first occurence (default
  `FALSE`).

## Value

A `<layout_info>` object, i.e. a list with the entries `index`,
`layout_name`, `layout_file`, `master_name`, `master_file`, and
`slide_layout`.
