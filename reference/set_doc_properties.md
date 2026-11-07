# Set document properties

set Word, PowerPoint or Excel document properties. These are not visible
in the document but are available as metadata of the document.

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

  an rdocx, rpptx or rxlsx object

- title, subject, creator, description:

  text fields

- created:

  a date-time object, written in UTC

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

The "modified" and "lastModifiedBy" properties are set when the document
is written: the current time, in UTC, and the name of the user
(`Sys.info()[["user"]]`). A document created from the default template,
with
[`read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.md),
[`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md)
or
[`read_xlsx()`](https://davidgohel.github.io/officer/reference/read_xlsx.md)
called without `path`, gets the current time as "created" property and
no "creator".

To write a fixed date in "created" and "modified", set option
`officer_datetime` to a date-time:
`options(officer_datetime = as.POSIXct("2026-01-01", tz = "UTC"))`. The
same code then produces the same file, byte for byte, whatever the time
zone or the locale, which makes it possible to compare files with a
checksum.

## See also

Other functions for Word document informations:
[`doc_properties()`](https://davidgohel.github.io/officer/reference/doc_properties.md),
[`docx_bookmarks()`](https://davidgohel.github.io/officer/reference/docx_bookmarks.md),
[`docx_dim()`](https://davidgohel.github.io/officer/reference/docx_dim.md),
[`length.rdocx()`](https://davidgohel.github.io/officer/reference/length.rdocx.md),
[`styles_info()`](https://davidgohel.github.io/officer/reference/styles_info.md)

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
