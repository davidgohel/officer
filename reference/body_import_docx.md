# Import an external docx in a 'Word' document

Import body content and footnotes of a Word document into an rdocx
object.

The function is similar to
[`body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.md)
but instead of adding the content as an external object, the document is
read and all its content is appended to the target document.

## Usage

``` r
body_import_docx(
  x,
  src,
  par_style_mapping = list(),
  run_style_mapping = list(),
  tbl_style_mapping = list(),
  prepend_chunks_on_styles = list()
)
```

## Arguments

- x:

  an rdocx object

- src:

  path to docx file to import

- par_style_mapping, run_style_mapping, tbl_style_mapping:

  Named lists describing how to remap styles from the source document
  (`src`) to styles available in the target document `x`. For each list
  entry, the name of the element is the target style (in `x`), and the
  value is a character vector of style names from the source document
  that should be replaced by this target style.

  - `par_style_mapping`: applies to paragraph styles.

  - `run_style_mapping`: applies to character (run) styles.

  - `tbl_style_mapping`: applies to table styles.

  Examples:

      par_style_mapping = list(
        "Normal"    = c("List Paragraph", "Body Text"),
        "heading 1" = "Heading 1"
      )
      run_style_mapping = list(
        "Emphasis"  = c("Emphasis", "Italic")
      )
      tbl_style_mapping = list(
        "Normal Table" = c("Light Shading")
      )

  Use
  [`styles_info()`](https://davidgohel.github.io/officer/reference/styles_info.md)
  to inspect available styles and verify their names.

- prepend_chunks_on_styles:

  A named list of run chunks to prepend to runs with specific styles.
  The names of the list are paragraph style names and the values run
  chunks to prepend. The first motivation for this argument is to allow
  prepending of runs in paragraphs with a defined style, for example to
  add a
  [`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md)
  with all image captions.

## Details

The following operations are performed when importing a document:

- Numberings are copied from the source document to the target document.

- Styles are not copied. If styles in the source document do not exist
  in the target document, the style specified in the
  `par_style_mapping`, `run_style_mapping` and `tbl_style_mapping`
  arguments will be used instead. If no mapping is provided, the default
  style will be used and a warning is emitted.

## See also

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/reference/body_add_blocks.md),
[`body_add_break()`](https://davidgohel.github.io/officer/reference/body_add_break.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/reference/body_add_caption.md),
[`body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/reference/body_add_fpar.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/reference/body_add_gg.md),
[`body_add_img()`](https://davidgohel.github.io/officer/reference/body_add_img.md),
[`body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/reference/body_add_plot.md),
[`body_add_table()`](https://davidgohel.github.io/officer/reference/body_add_table.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/reference/body_append_context.md)

## Examples

``` r
library(officer)

# example file from the package
file_input <- system.file(
  package = "officer",
  "doc_examples/example.docx"
)

# create a new rdocx document
x <- read_docx()

# import content from file_input
x <- body_import_docx(
  x = x,
  src = file_input,
  # style mapping for paragraphs and tables
  par_style_mapping = list(
    "Normal" = c("List Paragraph")
  ),
  tbl_style_mapping = list(
    "Normal Table" = "Light Shading"
  )
)

# Create temporary file
tf <- tempfile(fileext = ".docx")
# write to file
print(x, target = tf)
```
