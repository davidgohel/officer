# Detect and handle duplicate placeholder labels

PowerPoint does not enforce unique placeholder labels in a layout.
Selecting a placeholder via its label using
[ph_location_label](https://davidgohel.github.io/officer/dev/reference/ph_location_label.md)
will throw an error, if the label is not unique. layout_dedupe_ph_labels
helps to detect, rename, or delete duplicate placholder labels.

## Usage

``` r
layout_dedupe_ph_labels(x, action = "detect", print_info = FALSE)
```

## Arguments

- x:

  An `rpptx` object.

- action:

  Action to perform on duplicate placeholder labels. One of:

  - `detect` (default) = show info on dupes only, make no changes

  - `rename` = create unique labels. Labels are renamed by appending a
    sequential number separated by dot to duplicate labels. For example,
    `c("title", "title")` becomes `c("title.1", "title.2")`.

  - `delete` = only keep one of the placeholders with a duplicate label

- print_info:

  Print action information (e.g. renamed placeholders) to console?
  Default is `FALSE`. Always `TRUE` for action `detect`.

## Value

A `rpptx` object (with modified placeholder labels).

## Examples

``` r
x <- read_pptx()
layout_dedupe_ph_labels(x)
#> No duplicate placeholder labels detected.

file <- system.file("doc_examples", "ph_dupes.pptx", package = "officer")
x <- read_pptx(file)
layout_dedupe_ph_labels(x)
#> Placeholders with duplicate labels:
#> * 'ph_label_new' = new placeholder label for action = 'rename'
#> * 'delete_flag' = deleted placeholders for action = 'delete'
#>   master_name layout_name  ph_label ph_label_new delete_flag
#> 1     Master1     2-dupes Content 7  Content 7.1       FALSE
#> 2     Master1     2-dupes Content 7  Content 7.2        TRUE
layout_dedupe_ph_labels(x, "rename", print_info = TRUE)
#> Renamed duplicate placeholder labels:
#> * 'ph_label_new' = new placeholder label
#>   master_name layout_name  ph_label ph_label_new
#> 1     Master1     2-dupes Content 7  Content 7.1
#> 2     Master1     2-dupes Content 7  Content 7.2
```
