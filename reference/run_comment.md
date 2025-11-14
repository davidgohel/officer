# Comment for 'Word'

Add a comment on a run object.

## Usage

``` r
run_comment(
  cmt,
  run = ftext(""),
  author = "",
  date = "",
  initials = "",
  prop = NULL
)
```

## Arguments

- cmt:

  a set of blocks to be used as comment content returned by function
  [`block_list()`](https://davidgohel.github.io/officer/reference/block_list.md).
  the "run functions for reporting".

- run:

  a run object, made with a call to one of

- author:

  comment author.

- date:

  comment date

- initials:

  comment initials

- prop:

  formatting text properties returned by
  [`fp_text_lite()`](https://davidgohel.github.io/officer/reference/fp_text.md)
  or
  [`fp_text()`](https://davidgohel.github.io/officer/reference/fp_text.md).
  It also can be NULL in which case, no formatting is defined (the
  default is applied).

## See also

Other run functions for reporting:
[`external_img()`](https://davidgohel.github.io/officer/reference/external_img.md),
[`floating_external_img()`](https://davidgohel.github.io/officer/reference/floating_external_img.md),
[`ftext()`](https://davidgohel.github.io/officer/reference/ftext.md),
[`hyperlink_ftext()`](https://davidgohel.github.io/officer/reference/hyperlink_ftext.md),
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md),
[`run_bookmark()`](https://davidgohel.github.io/officer/reference/run_bookmark.md),
[`run_columnbreak()`](https://davidgohel.github.io/officer/reference/run_columnbreak.md),
[`run_footnote()`](https://davidgohel.github.io/officer/reference/run_footnote.md),
[`run_footnoteref()`](https://davidgohel.github.io/officer/reference/run_footnoteref.md),
[`run_linebreak()`](https://davidgohel.github.io/officer/reference/run_linebreak.md),
[`run_pagebreak()`](https://davidgohel.github.io/officer/reference/run_pagebreak.md),
[`run_reference()`](https://davidgohel.github.io/officer/reference/run_reference.md),
[`run_tab()`](https://davidgohel.github.io/officer/reference/run_tab.md),
[`run_word_field()`](https://davidgohel.github.io/officer/reference/run_word_field.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/reference/run_wordtext.md)

## Examples

``` r
fp_bold <- fp_text_lite(bold = TRUE)
fp_red <- fp_text_lite(color = "red")

bl <- block_list(
  fpar(ftext("Comment multiple words.", fp_bold)),
  fpar(
    ftext("Second line.", fp_red)
  )
)

comment1 <- run_comment(
  cmt = bl,
  run = ftext("with a comment"),
  author = "Author Me",
  date = Sys.Date(),
  initials = "AM"
)
par1 <- fpar("A paragraph ", comment1)

bl <- block_list(
  fpar(ftext("Comment a paragraph."))
)

comment2 <- run_comment(
  cmt = bl, run = ftext("A commented paragraph"),
  author = "Author You",
  date = Sys.Date(),
  initials = "AY"
)
par2 <- fpar(comment2)

doc <- read_docx()
doc <- body_add_fpar(doc, value = par1, style = "Normal")
doc <- body_add_fpar(doc, value = par2, style = "Normal")

print(doc, target = tempfile(fileext = ".docx"))
```
