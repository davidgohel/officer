# Add bookmark in a 'Word' document

Add a bookmark at the cursor location. The bookmark is added on the
first run of text in the current paragraph.

## Usage

``` r
body_bookmark(x, id)
```

## Arguments

- x:

  an rdocx object

- id:

  bookmark name

## Examples

``` r
# cursor_bookmark ----

doc <- read_docx()
doc <- body_add_par(doc, "centered text", style = "centered")
doc <- body_bookmark(doc, "text_to_replace")
```
