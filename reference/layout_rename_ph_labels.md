# Change ph labels in a layout

There are two versions of the function. The first takes a set of
key-value pairs to rename the ph labels. The second uses a right hand
side (rhs) assignment to specify the new ph labels. See section
*Details*.  
  
*NB:* You can also rename ph labels directly in PowerPoint. Open the
master template view (`Alt` + `F10`) and go to `Home` \> `Arrange` \>
`Selection Pane`.

## Usage

``` r
layout_rename_ph_labels(x, layout, master = NULL, ..., .dots = NULL)

layout_rename_ph_labels(x, layout, master = NULL, id = NULL) <- value
```

## Arguments

- x:

  An `rpptx` object.

- layout:

  Layout name or index. Index is the row index of
  [`layout_summary()`](https://davidgohel.github.io/officer/reference/layout_summary.md).

- master:

  Name of master. Only required if the layout name is not unique across
  masters.

- ...:

  Comma separated list of key-value pairs to rename phs. Either
  reference a ph via its label (`"old label"` = `"new label"`) or its
  unique id (`"id"` = `"new label"`).

- .dots:

  Provide a named list or vector of key-value pairs to rename phs
  (`list("old label"` = `"new label"`).

- id:

  Unique placeholder id (see column `id` in
  [`layout_properties()`](https://davidgohel.github.io/officer/reference/layout_properties.md)
  or
  [`plot_layout_properties()`](https://davidgohel.github.io/officer/reference/plot_layout_properties.md)).

- value:

  Not relevant for user. A pure technical necessity for rhs assignments.

## Value

Vector of renamed ph labels.

## Details

- Note the difference between the terms `id` and `index`. Both can be
  found in the output of
  [`layout_properties()`](https://davidgohel.github.io/officer/reference/layout_properties.md).
  The unique ph `id` is found in column `id`. The `index` refers to the
  index of the data frame row.

- In a right hand side (rhs) label assignment (`<- new_labels`), there
  are two ways to optionally specify a subset of phs to rename. In both
  cases, the length of the rhs vector (the new labels) must match the
  length of the id or index:

  1.  use the `id` argument to specify ph ids to rename:
      `layout_rename_ph_labels(..., id = 2:3) <- new_labels`

  2.  use an `index` in squared brackets:
      `layout_rename_ph_labels(...)[1:2] <- new_labels`

## Examples

``` r
x <- read_pptx()

# INFO -------------

# Returns layout's ph_labels by default in same order as layout_properties()
layout_rename_ph_labels(x, "Comparison")
#> [1] "Title 1"                    "Text Placeholder 2"        
#> [3] "Text Placeholder 4"         "Content Placeholder 3"     
#> [5] "Content Placeholder 5"      "Date Placeholder 6"        
#> [7] "Footer Placeholder 7"       "Slide Number Placeholder 8"
layout_properties(x, "Comparison")$ph_label
#> [1] "Title 1"                    "Text Placeholder 2"        
#> [3] "Text Placeholder 4"         "Content Placeholder 3"     
#> [5] "Content Placeholder 5"      "Date Placeholder 6"        
#> [7] "Footer Placeholder 7"       "Slide Number Placeholder 8"


# BASICS -----------
#
# HINT: run `plot_layout_properties(x, "Comparison")` to see how labels change

# rename using key-value pairs: 'old label' = 'new label' or 'id' = 'new label'
layout_rename_ph_labels(x, "Comparison", "Title 1" = "LABEL MATCHED") # label matching
layout_rename_ph_labels(x, "Comparison", "3" = "ID MATCHED") # id matching
layout_rename_ph_labels(
  x,
  "Comparison",
  "Date Placeholder 6" = "DATE",
  "8" = "FOOTER"
) # label, id

# rename using a named list and the .dots arg
renames <- list("Content Placeholder 3" = "CONTENT_1", "6" = "CONTENT_2")
layout_rename_ph_labels(x, "Comparison", .dots = renames)

# rename via rhs assignment and optional index (not id!)
layout_rename_ph_labels(x, "Comparison") <- LETTERS[1:8]
layout_rename_ph_labels(x, "Comparison")[1:3] <- paste("CHANGED", 1:3)

# rename via rhs assignment and ph id (not index)
layout_rename_ph_labels(x, "Comparison", id = c(2, 4)) <- paste("ID =", c(2, 4))


# MORE ------------

# make all labels lower case
labels <- layout_rename_ph_labels(x, "Comparison")
layout_rename_ph_labels(x, "Comparison") <- tolower(labels)

# rename all labels to type [type_idx]
lp <- layout_properties(x, "Comparison")
layout_rename_ph_labels(x, "Comparison") <- paste0(
  lp$type,
  " [",
  lp$type_idx,
  "]"
)

# rename duplicated placeholders (see also `layout_dedupe_ph_labels()`)
file <- system.file("doc_examples", "ph_dupes.pptx", package = "officer")
x <- read_pptx(file)
lp <- layout_properties(x, "2-dupes")
idx <- which(lp$ph_label == "Content 7") # exists twice
layout_rename_ph_labels(x, "2-dupes")[idx] <- paste("DUPLICATE", seq_along(idx))

# warning: in case of duped labels only the first occurrence is renamed
x <- read_pptx(file)
layout_rename_ph_labels(x, "2-dupes", "Content 7" = "new label")
#> Warning: When renaming a label with duplicates, only the first occurrence is renamed.
#> ✖ Renaming 1 ph label with duplicates: "Content 7"
```
