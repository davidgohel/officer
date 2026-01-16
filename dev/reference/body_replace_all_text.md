# Replace text anywhere in the document

Replace text anywhere in the document, or at a cursor.

Replace all occurrences of old_value with new_value. This method uses
[`grepl()`](https://rdrr.io/r/base/grep.html)/[`gsub()`](https://rdrr.io/r/base/grep.html)
for pattern matching; you may supply arguments as required (and
therefore use [`regex()`](https://rdrr.io/r/base/regex.html) features)
using the optional `...` argument.

Note that by default, grepl/gsub will use `fixed=FALSE`, which means
that `old_value` and `new_value` will be interepreted as regular
expressions.

**Chunking of text**

Note that the behind-the-scenes representation of text in a Word
document is frequently not what you might expect! Sometimes a paragraph
of text is broken up (or "chunked") into several "runs," as a result of
style changes, pauses in text entry, later revisions and edits, etc. If
you have not styled the text, and have entered it in an "all-at-once"
fashion, e.g. by pasting it or by outputing it programmatically into
your Word document, then this will likely not be a problem. If you are
working with a manually-edited document, however, this can lead to
unexpected failures to find text.

You can use the officer function
[`docx_show_chunk()`](https://davidgohel.github.io/officer/dev/reference/docx_show_chunk.md)
to show how the paragraph of text at the current cursor has been chunked
into runs, and what text is in each chunk. This can help troubleshoot
unexpected failures to find text.

## Usage

``` r
body_replace_all_text(
  x,
  old_value,
  new_value,
  only_at_cursor = FALSE,
  warn = TRUE,
  ...
)

headers_replace_all_text(
  x,
  old_value,
  new_value,
  only_at_cursor = FALSE,
  warn = TRUE,
  ...
)

footers_replace_all_text(
  x,
  old_value,
  new_value,
  only_at_cursor = FALSE,
  warn = TRUE,
  ...
)
```

## Arguments

- x:

  a docx device

- old_value:

  the value to replace

- new_value:

  the value to replace it with

- only_at_cursor:

  if `TRUE`, only search-and-replace at the current cursor; if `FALSE`
  (default), search-and-replace in the entire document (this can be slow
  on large documents!)

- warn:

  warn if `old_value` could not be found.

- ...:

  optional arguments to grepl/gsub (e.g. `fixed=TRUE`)

## header_replace_all_text

Replacements will be performed in each header of all sections.

Replacements will be performed in each footer of all sections.

## See also

[`grepl()`](https://rdrr.io/r/base/grep.html),
[`regex()`](https://rdrr.io/r/base/regex.html),
[`docx_show_chunk()`](https://davidgohel.github.io/officer/dev/reference/docx_show_chunk.md)

## Author

Frank Hangler, <frank@plotandscatter.com>

## Examples

``` r
library(officer)

doc <- read_docx()
doc <- body_add_par(doc, "Placeholder one")
doc <- body_add_par(doc, "Placeholder two")

# Show text chunk at cursor
docx_show_chunk(doc) # Output is 'Placeholder two'
#> 1 text nodes found at this cursor. 
#>   <w:t>: 'Placeholder two'

# Simple search-and-replace at current cursor, with regex turned off
doc <- body_replace_all_text(
  doc,
  old_value = "Placeholder",
  new_value = "new",
  only_at_cursor = TRUE,
  fixed = TRUE
)
docx_show_chunk(doc) # Output is 'new two'
#> 1 text nodes found at this cursor. 
#>   <w:t>: 'new two'

# Do the same, but in the entire document and ignoring case
doc <- body_replace_all_text(
  doc,
  old_value = "placeholder",
  new_value = "new",
  only_at_cursor = FALSE,
  ignore.case = TRUE
)
doc <- cursor_backward(doc)
docx_show_chunk(doc) # Output is 'new one'
#> 1 text nodes found at this cursor. 
#>   <w:t>: 'new one'

# Use regex : replace all words starting with "n" with the word "example"
doc <- body_replace_all_text(doc, "\\bn.*?\\b", "example")
docx_show_chunk(doc) # Output is 'example one'
#> 1 text nodes found at this cursor. 
#>   <w:t>: 'example one'
```
