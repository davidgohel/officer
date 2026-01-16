# Set document properties

set Word or PowerPoint document properties. These are not visible in the
document but are available as metadata of the document.

Any character property can be added as a document property. It provides
an easy way to insert arbitrary fields. Given the challenges that can be
encountered with find-and-replace in word with officer, the use of
document fields and quick text fields provides a much more robust
approach to automatic document generation from R.

## Usage

``` r
set_doc_properties(
  x,
  title = NULL,
  subject = NULL,
  creator = NULL,
  description = NULL,
  created = NULL,
  hyperlink_base = NULL,
  ...,
  values = NULL
)
```

## Arguments

- x:

  an rdocx or rpptx object

- title, subject, creator, description:

  text fields

- created:

  a date object

- hyperlink_base:

  a string specifying the base URL for relative hyperlinks in the
  document (only for rdocx).

- ...:

  named arguments (names are field names), each element is a single
  character value specifying value associated with the corresponding
  field name. These pairs of *key-value* are added as custom properties.
  If a value is `NULL` or `NA`, the corresponding field is set to ” in
  the document properties.

- values:

  a named list (names are field names), each element is a single
  character value specifying value associated with the corresponding
  field name. If `values` is provided, argument `...` will be ignored.

## Note

The "last modified" and "last modified by" fields will be automatically
be updated when the file is written.

## See also

Other functions for Word document informations:
[`doc_properties()`](https://davidgohel.github.io/officer/dev/reference/doc_properties.md),
[`docx_bookmarks()`](https://davidgohel.github.io/officer/dev/reference/docx_bookmarks.md),
[`docx_dim()`](https://davidgohel.github.io/officer/dev/reference/docx_dim.md),
[`length.rdocx()`](https://davidgohel.github.io/officer/dev/reference/length.rdocx.md),
[`styles_info()`](https://davidgohel.github.io/officer/dev/reference/styles_info.md)

## Examples

``` r
x <- read_docx()
x <- set_doc_properties(x, title = "title",
  subject = "document subject", creator = "Me me me",
  description = "this document is empty",
  created = Sys.time(),
  yoyo = "yok yok",
  glop = "pas glop")
x <- doc_properties(x)
```
