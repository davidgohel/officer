# Set cursor in a 'Word' document

A set of functions is available to manipulate the position of a virtual
cursor. This cursor will be used when inserting, deleting or updating
elements in the document.

## Usage

``` r
cursor_begin(x)

cursor_bookmark(x, id)

cursor_end(x)

cursor_reach_index(x, index)

cursor_reach(x, keyword, fixed = FALSE)

cursor_reach_test(x, keyword)

cursor_forward(x)

cursor_backward(x)
```

## Arguments

- x:

  a docx device

- id:

  bookmark id

- index:

  element index in the document

- keyword:

  keyword to look for as a regular expression

- fixed:

  logical. If TRUE, pattern is a string to be matched as is.

## cursor_begin

Set the cursor at the beginning of the document, on the first element of
the document (usually a paragraph or a table).

## cursor_bookmark

Set the cursor at a bookmark that has previously been set.

## cursor_end

Set the cursor at the end of the document, on the last element of the
document.

## cursor_reach_index

Set the cursor at a specific index position in the document.

## cursor_reach

Set the cursor on the first element of the document that contains text
specified in argument `keyword`. The argument `keyword` is a regexpr
pattern.

## cursor_reach_test

Test if an expression has a match in the document that contains text
specified in argument `keyword`. The argument `keyword` is a regexpr
pattern.

## cursor_forward

Move the cursor forward, it increments the cursor in the document.

## cursor_backward

Move the cursor backward, it decrements the cursor in the document.

## Examples

``` r
library(officer)

# create a template ----
doc <- read_docx()
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "Hello text to replace")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "blah blah blah")
doc <- body_add_par(doc, "Hello text to replace")
doc <- body_add_par(doc, "blah blah blah")
template_file <- print(
  x = doc,
  target = tempfile(fileext = ".docx")
)

# replace all pars containing "to replace" ----
doc <- read_docx(path = template_file)
while (cursor_reach_test(doc, "to replace")) {
  doc <- cursor_reach(doc, "to replace")

  doc <- body_add_fpar(
    x = doc,
    pos = "on",
    value = fpar(
      "Here is a link: ",
      hyperlink_ftext(
        text = "yopyop",
        href = "https://cran.r-project.org/"
      )
    )
  )
}

doc <- cursor_end(doc)
doc <- body_add_par(doc, "Yap yap yap yap...")

result_file <- print(
  x = doc,
  target = tempfile(fileext = ".docx")
)

# cursor_bookmark ----

doc <- read_docx()
doc <- body_add_par(doc, "centered text", style = "centered")
doc <- body_bookmark(doc, "text_to_replace")
doc <- body_add_par(doc, "A title", style = "heading 1")
doc <- body_add_par(doc, "Hello world!", style = "Normal")
doc <- cursor_bookmark(doc, "text_to_replace")
doc <- body_add_table(doc, value = iris, style = "table_template")

print(doc, target = tempfile(fileext = ".docx"))
```
