# Word footnote reference

Wraps a footnote reference in an object that can then be inserted as a
run/chunk with
[`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md)
or within an R Markdown document.

## Usage

``` r
run_footnoteref(prop = NULL)
```

## Arguments

- prop:

  formatting text properties returned by
  [`fp_text_lite()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)
  or
  [`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md).
  It also can be NULL in which case, no formatting is defined (the
  default is applied).

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
[`run_linebreak()`](https://davidgohel.github.io/officer/dev/reference/run_linebreak.md),
[`run_pagebreak()`](https://davidgohel.github.io/officer/dev/reference/run_pagebreak.md),
[`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md),
[`run_tab()`](https://davidgohel.github.io/officer/dev/reference/run_tab.md),
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/dev/reference/run_wordtext.md)

## Examples

``` r
run_footnoteref()
#> $pr
#> NULL
#> 
#> attr(,"class")
#> [1] "run_footnoteref" "run"            
to_wml(run_footnoteref())
#> [1] "<w:r><w:footnoteRef/></w:r>"
```
