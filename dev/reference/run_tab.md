# Tab for 'Word'

Object representing a tab in a Word document. The result must be used
within a call to
[fpar](https://davidgohel.github.io/officer/dev/reference/fpar.md). It
will only have effects in Word output.

Tabulation marks settings can be defined with
[`fp_tabs()`](https://davidgohel.github.io/officer/dev/reference/fp_tabs.md)
in paragraph settings defined with
[`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md).

## Usage

``` r
run_tab()
```

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
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/dev/reference/run_wordtext.md)

## Examples

``` r
z <- fp_tabs(
  fp_tab(pos = 0.5, style = "decimal"),
  fp_tab(pos = 1.5, style = "decimal")
)
par1 <- fpar(
  run_tab(), ftext("88."),
  run_tab(), ftext("987.45"),
  fp_p = fp_par(
    tabs = z
  )
)
par2 <- fpar(
  run_tab(), ftext("8."),
  run_tab(), ftext("670987.45"),
  fp_p = fp_par(
    tabs = z
  )
)
x <- read_docx()
x <- body_add(x, par1)
x <- body_add(x, par2)
print(x, target = tempfile(fileext = ".docx"))
```
