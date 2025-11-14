# Write a 'Word' File

`print.rdocx()` is the essential output function for creating Word files
with officer. It takes an `rdocx` object (created with
[`read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.md)
and populated with content) and writes it to disk as a `.docx` file.

This function performs all necessary post-processing operations before
writing the file.

The function is typically called at the end of your document creation
workflow, after all content has been added with `body_add_*()`
functions.

## Usage

``` r
# S3 method for class 'rdocx'
print(
  x,
  target = NULL,
  copy_header_refs = FALSE,
  copy_footer_refs = FALSE,
  preview = FALSE,
  ...
)
```

## Arguments

- x:

  an `rdocx` object created with
  [`read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.md)

- target:

  path to the `.docx` file to write. The file will be created or
  overwritten if it already exists. If `NULL` and `preview = FALSE`, the
  function returns `NULL` without writing a file.

- copy_header_refs, copy_footer_refs:

  logical, default is FALSE. If TRUE, copy the references to the header
  and footer in each section of the body of the document. This parameter
  is experimental and may change in a future version.

- preview:

  Save `x` to a temporary file and open it (default `FALSE`). When
  `TRUE`, the document is saved to a temporary location and opened with
  the system's default application for `.docx` files, useful for quick
  previewing during development.

- ...:

  unused

## Value

The full path to the created `.docx` file (invisibly). This allows
chaining operations or capturing the output path for further use.

## See also

Create a 'Word' document object with
[`read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.md),
add content with functions
[`body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/reference/body_add_plot.md),
[`body_add_table()`](https://davidgohel.github.io/officer/reference/body_add_table.md),
change settings with
[`docx_set_settings()`](https://davidgohel.github.io/officer/reference/docx_set_settings.md),
set properties with
[`set_doc_properties()`](https://davidgohel.github.io/officer/reference/set_doc_properties.md),
read 'Word' styles with
[`styles_info()`](https://davidgohel.github.io/officer/reference/styles_info.md).

## Examples

``` r
library(officer)

# This example demonstrates how to create
# an small document -----

## Create a new Word document
doc <- read_docx()
doc <- body_add_par(doc, "hello world")
## Save the document
output_file <- print(doc, target = tempfile(fileext = ".docx"))

# preview mode: save to temp file and open locally ----
if (FALSE) { # \dontrun{
print(doc, preview = TRUE)
} # }
```
