# Auto number

Create an autonumbered chunk, i.e. a string representation of a
sequence, each item will be numbered. These runs can also be bookmarked
and be used later for cross references.

## Usage

``` r
run_autonum(
  seq_id = "table",
  pre_label = "Table ",
  post_label = ": ",
  bkm = NULL,
  bkm_all = FALSE,
  prop = NULL,
  start_at = NULL,
  tnd = 0,
  tns = "-"
)
```

## Arguments

- seq_id:

  sequence identifier

- pre_label, post_label:

  text to add before and after number

- bkm:

  bookmark id to associate with autonumber run. If NULL, no bookmark is
  added. Value can only be made of alpha numeric characters, ':', -' and
  '\_'.

- bkm_all:

  if TRUE, the bookmark will be set on the whole string, if FALSE, the
  bookmark will be set on the number only. Default to FALSE. As an
  effect when a reference to this bookmark is used, the text can be like
  "Table 1" or "1" (pre_label is not included in the referenced text).

- prop:

  formatting text properties returned by
  [fp_text](https://davidgohel.github.io/officer/dev/reference/fp_text.md).

- start_at:

  If not NULL, it must be a positive integer, it specifies the new
  number to use, at which number the auto numbering will restart.

- tnd:

  *title number depth*, a positive integer (only applies if positive)
  that specify the depth (or heading of level *depth*) to use for
  prefixing the caption number with this last reference number. For
  example, setting `tnd=2` will generate numbered captions like '4.3-2'
  (figure 2 of chapter 4.3).

- tns:

  separator to use between title number and table number. Default is
  "-".

## usage

You can use this function in conjunction with
[fpar](https://davidgohel.github.io/officer/dev/reference/fpar.md) to
create paragraphs consisting of differently formatted text parts. You
can also use this function as an *r chunk* in an R Markdown document
made with package officedown.

## See also

Other run functions for reporting:
[`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md),
[`floating_external_img()`](https://davidgohel.github.io/officer/dev/reference/floating_external_img.md),
[`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md),
[`hyperlink_ftext()`](https://davidgohel.github.io/officer/dev/reference/hyperlink_ftext.md),
[`run_bookmark()`](https://davidgohel.github.io/officer/dev/reference/run_bookmark.md),
[`run_columnbreak()`](https://davidgohel.github.io/officer/dev/reference/run_columnbreak.md),
[`run_comment()`](https://davidgohel.github.io/officer/dev/reference/run_comment.md),
[`run_footnote()`](https://davidgohel.github.io/officer/dev/reference/run_footnote.md),
[`run_footnoteref()`](https://davidgohel.github.io/officer/dev/reference/run_footnoteref.md),
[`run_linebreak()`](https://davidgohel.github.io/officer/dev/reference/run_linebreak.md),
[`run_pagebreak()`](https://davidgohel.github.io/officer/dev/reference/run_pagebreak.md),
[`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md),
[`run_tab()`](https://davidgohel.github.io/officer/dev/reference/run_tab.md),
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/dev/reference/run_wordtext.md)

Other Word computed fields:
[`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md),
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md)

## Examples

``` r
run_autonum()
#> $seq_id
#> [1] "table"
#> 
#> $pre_label
#> [1] "Table "
#> 
#> $post_label
#> [1] ": "
#> 
#> $bookmark
#> NULL
#> 
#> $bookmark_all
#> [1] FALSE
#> 
#> $pr
#> NULL
#> 
#> $start_at
#> NULL
#> 
#> $tnd
#> [1] 0
#> 
#> $tns
#> [1] "-"
#> 
#> attr(,"class")
#> [1] "run_autonum" "run"        
run_autonum(seq_id = "fig", pre_label = "fig. ")
#> $seq_id
#> [1] "fig"
#> 
#> $pre_label
#> [1] "fig. "
#> 
#> $post_label
#> [1] ": "
#> 
#> $bookmark
#> NULL
#> 
#> $bookmark_all
#> [1] FALSE
#> 
#> $pr
#> NULL
#> 
#> $start_at
#> NULL
#> 
#> $tnd
#> [1] 0
#> 
#> $tns
#> [1] "-"
#> 
#> attr(,"class")
#> [1] "run_autonum" "run"        
run_autonum(seq_id = "tab", pre_label = "Table ", bkm = "anytable")
#> $seq_id
#> [1] "tab"
#> 
#> $pre_label
#> [1] "Table "
#> 
#> $post_label
#> [1] ": "
#> 
#> $bookmark
#> [1] "anytable"
#> 
#> $bookmark_all
#> [1] FALSE
#> 
#> $pr
#> NULL
#> 
#> $start_at
#> NULL
#> 
#> $tnd
#> [1] 0
#> 
#> $tns
#> [1] "-"
#> 
#> attr(,"class")
#> [1] "run_autonum" "run"        
run_autonum(
  seq_id = "tab", pre_label = "Table ", bkm = "anytable",
  tnd = 2, tns = " "
)
#> $seq_id
#> [1] "tab"
#> 
#> $pre_label
#> [1] "Table "
#> 
#> $post_label
#> [1] ": "
#> 
#> $bookmark
#> [1] "anytable"
#> 
#> $bookmark_all
#> [1] FALSE
#> 
#> $pr
#> NULL
#> 
#> $start_at
#> NULL
#> 
#> $tnd
#> [1] 2
#> 
#> $tns
#> [1] " "
#> 
#> attr(,"class")
#> [1] "run_autonum" "run"        
```
