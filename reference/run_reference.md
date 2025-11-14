# Cross reference

Create a representation of a reference

## Usage

``` r
run_reference(id, prop = NULL)
```

## Arguments

- id:

  reference id, a string

- prop:

  formatting text properties returned by
  [fp_text](https://davidgohel.github.io/officer/reference/fp_text.md).

## usage

You can use this function in conjunction with
[fpar](https://davidgohel.github.io/officer/reference/fpar.md) to create
paragraphs consisting of differently formatted text parts. You can also
use this function as an *r chunk* in an R Markdown document made with
package officedown.

## See also

Other run functions for reporting:
[`external_img()`](https://davidgohel.github.io/officer/reference/external_img.md),
[`floating_external_img()`](https://davidgohel.github.io/officer/reference/floating_external_img.md),
[`ftext()`](https://davidgohel.github.io/officer/reference/ftext.md),
[`hyperlink_ftext()`](https://davidgohel.github.io/officer/reference/hyperlink_ftext.md),
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md),
[`run_bookmark()`](https://davidgohel.github.io/officer/reference/run_bookmark.md),
[`run_columnbreak()`](https://davidgohel.github.io/officer/reference/run_columnbreak.md),
[`run_comment()`](https://davidgohel.github.io/officer/reference/run_comment.md),
[`run_footnote()`](https://davidgohel.github.io/officer/reference/run_footnote.md),
[`run_footnoteref()`](https://davidgohel.github.io/officer/reference/run_footnoteref.md),
[`run_linebreak()`](https://davidgohel.github.io/officer/reference/run_linebreak.md),
[`run_pagebreak()`](https://davidgohel.github.io/officer/reference/run_pagebreak.md),
[`run_tab()`](https://davidgohel.github.io/officer/reference/run_tab.md),
[`run_word_field()`](https://davidgohel.github.io/officer/reference/run_word_field.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/reference/run_wordtext.md)

Other Word computed fields:
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md),
[`run_word_field()`](https://davidgohel.github.io/officer/reference/run_word_field.md)

## Examples

``` r
run_reference("a_ref")
#> $id
#> [1] "a_ref"
#> 
#> $pr
#> NULL
#> 
#> attr(,"class")
#> [1] "run_reference" "run"          
```
