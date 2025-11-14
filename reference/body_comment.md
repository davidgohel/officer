# Add comment in a 'Word' document

Add a comment at the cursor location. The comment is added on the first
run of text in the current paragraph.

## Usage

``` r
body_comment(x, cmt = ftext(""), author = "", date = "", initials = "")
```

## Arguments

- x:

  an rdocx object

- cmt:

  a set of blocks to be used as comment content returned by function
  [`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md).

- author:

  comment author.

- date:

  comment date

- initials:

  comment initials

## Examples

``` r
doc <- read_docx()
doc <- body_add_par(doc, "Paragraph")
doc <- body_comment(doc, block_list("This is a comment."))
docx_file <- print(doc, target = tempfile(fileext = ".docx"))
docx_comments(read_docx(docx_file))
#>   comment_id author initials date               text para_id commented_text
#> 1          0                      This is a comment.      NA      Paragraph
```
