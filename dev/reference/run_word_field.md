# 'Word' computed field

Create a 'Word' computed field.

## Usage

``` r
run_word_field(field, prop = NULL, seqfield = NULL)

run_seqfield(field, prop = NULL, seqfield = NULL)
```

## Arguments

- field:

  Value for a "Word Computed Field" as a string.

- prop:

  formatting text properties returned by
  [fp_text](https://davidgohel.github.io/officer/dev/reference/fp_text.md).

- seqfield:

  deprecated in favor of `field`.

## Note

In the previous version, this function was called `run_seqfield` but the
name was wrong and should have been `run_word_field`.

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
[`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md),
[`run_bookmark()`](https://davidgohel.github.io/officer/dev/reference/run_bookmark.md),
[`run_columnbreak()`](https://davidgohel.github.io/officer/dev/reference/run_columnbreak.md),
[`run_comment()`](https://davidgohel.github.io/officer/dev/reference/run_comment.md),
[`run_footnote()`](https://davidgohel.github.io/officer/dev/reference/run_footnote.md),
[`run_footnoteref()`](https://davidgohel.github.io/officer/dev/reference/run_footnoteref.md),
[`run_linebreak()`](https://davidgohel.github.io/officer/dev/reference/run_linebreak.md),
[`run_pagebreak()`](https://davidgohel.github.io/officer/dev/reference/run_pagebreak.md),
[`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md),
[`run_tab()`](https://davidgohel.github.io/officer/dev/reference/run_tab.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/dev/reference/run_wordtext.md)

Other Word computed fields:
[`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md),
[`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md)

## Examples

``` r
run_word_field(field = "PAGE  \\* MERGEFORMAT")
#> $field
#> [1] "PAGE  \\* MERGEFORMAT"
#> 
#> $pr
#> NULL
#> 
#> attr(,"class")
#> [1] "run_word_field" "run"           
run_word_field(field = "Date \\@ \"MMMM d yyyy\"")
#> $field
#> [1] "Date \\@ \"MMMM d yyyy\""
#> 
#> $pr
#> NULL
#> 
#> attr(,"class")
#> [1] "run_word_field" "run"           
```
