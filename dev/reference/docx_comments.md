# Get comments in a Word document as a data.frame

return a data.frame representing the comments in a Word document.

## Usage

``` r
docx_comments(x)
```

## Arguments

- x:

  an rdocx object

## Details

Each row of the returned data frame contains data for one comment. The
columns contain the following information:

- "comment_id" - unique comment id

- "author" - name of the comment author

- "initials" - initials of the comment author

- "date" - timestamp of the comment

- "text" - a list column of characters containing the comment text.
  Elements can be vectors of length \> 1 if a comment contains multiple
  paragraphs, blocks or runs or of length 0 if the comment is empty.

- "para_id" - a list column of characters containing the parent
  paragraph IDs. Elememts can be vectors of length \> 1 if a comment
  spans multiple paragraphs or of length 0 if the comment has no parent
  paragraph.

- "commented_text" - a list column of characters containing the
  commented text. Elements can be vectors of length \> 1 if a comment
  spans multiple paragraphs or runs or of length 0 if the commented text
  is empty.

## Examples

``` r
bl <- block_list(
  fpar("Comment multiple words."),
  fpar("Second line")
)

a_par <- fpar(
  "This paragraph contains",
  run_comment(
    cmt = bl,
    run = ftext("a comment."),
    author = "Author Me",
    date = "2023-06-01"
  )
)

doc <- read_docx()
doc <- body_add_fpar(doc, value = a_par, style = "Normal")

docx_file <- print(doc, target = tempfile(fileext = ".docx"))

docx_comments(read_docx(docx_file))
#>   comment_id    author initials       date                                 text
#> 1          0 Author Me          2023-06-01 Comment multiple words., Second line
#>   para_id commented_text
#> 1      NA     a comment.
```
