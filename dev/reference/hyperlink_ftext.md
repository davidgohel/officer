# Formatted chunk of text with hyperlink

Format a chunk of text with text formatting properties (bold, color,
...), the chunk is associated with an hyperlink.

## Usage

``` r
hyperlink_ftext(text, prop = NULL, href)
```

## Arguments

- text:

  text value, a single character value

- prop:

  formatting text properties returned by
  [fp_text](https://davidgohel.github.io/officer/dev/reference/fp_text.md).
  It also can be NULL in which case, no formatting is defined (the
  default is applied).

- href:

  URL value

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
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/dev/reference/run_wordtext.md)

## Examples

``` r
ft <- fp_text(font.size = 12, bold = TRUE)
hyperlink_ftext(
  href = "https://cran.r-project.org/index.html",
  text = "some text", prop = ft
)
#> $value
#> [1] "some text"
#> 
#> $pr
#>   font.size italic bold underlined strike color     shading fontname
#> 1        12  FALSE TRUE      FALSE  FALSE black transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> $href
#> [1] "https://cran.r-project.org/index.html"
#> 
#> attr(,"class")
#> [1] "hyperlink_ftext" "cot"             "run"            
```
