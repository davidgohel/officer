# Formatted chunk of text

Format a chunk of text with text formatting properties (bold, color,
...). The function allows you to create pieces of text formatted the way
you want.

## Usage

``` r
ftext(text, prop = NULL)
```

## Arguments

- text:

  text value, a single character value

- prop:

  formatting text properties returned by
  [fp_text](https://davidgohel.github.io/officer/reference/fp_text.md).
  It also can be NULL in which case, no formatting is defined (the
  default is applied).

## usage

You can use this function in conjunction with
[fpar](https://davidgohel.github.io/officer/reference/fpar.md) to create
paragraphs consisting of differently formatted text parts. You can also
use this function as an *r chunk* in an R Markdown document made with
package officedown.

## See also

[fp_text](https://davidgohel.github.io/officer/reference/fp_text.md)

Other run functions for reporting:
[`external_img()`](https://davidgohel.github.io/officer/reference/external_img.md),
[`floating_external_img()`](https://davidgohel.github.io/officer/reference/floating_external_img.md),
[`hyperlink_ftext()`](https://davidgohel.github.io/officer/reference/hyperlink_ftext.md),
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md),
[`run_bookmark()`](https://davidgohel.github.io/officer/reference/run_bookmark.md),
[`run_columnbreak()`](https://davidgohel.github.io/officer/reference/run_columnbreak.md),
[`run_comment()`](https://davidgohel.github.io/officer/reference/run_comment.md),
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
ftext("hello", fp_text())
#> text: hello
#> format:
#>   font.size italic  bold underlined strike color     shading fontname
#> 1        10  FALSE FALSE      FALSE  FALSE black transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline

properties1 <- fp_text(color = "red")
properties2 <- fp_text(bold = TRUE, shading.color = "yellow")
ftext1 <- ftext("hello", properties1)
ftext2 <- ftext("World", properties2)
paragraph <- fpar(ftext1, " ", ftext2)

x <- read_docx()
x <- body_add(x, paragraph)
print(x, target = tempfile(fileext = ".docx"))
```
