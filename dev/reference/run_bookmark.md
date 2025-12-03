# Bookmark for 'Word'

Add a bookmark on a run object.

## Usage

``` r
run_bookmark(bkm, run)
```

## Arguments

- bkm:

  bookmark id to associate with run. Value can only be made of alpha
  numeric characters, '-' and '\_'.

- run:

  a run object, made with a call to one of the "run functions for
  reporting".

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
run_bookmark("par1", ftext("some text", ft))
#> $id
#> [1] "par1"
#> 
#> $run
#> $run[[1]]
#> text: some text
#> format:
#>   font.size italic bold underlined strike color     shading fontname
#> 1        12  FALSE TRUE      FALSE  FALSE black transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> 
#> attr(,"class")
#> [1] "run_bookmark" "run"         
```
